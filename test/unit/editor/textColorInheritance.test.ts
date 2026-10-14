import { describe, expect, it } from 'vitest';
import { createEditableTextFixture } from '../../fixtures/editable-text';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { renderEditableText } from '../../../src/editor/renderEditableText';
import { renderSlide } from '../../../src/renderer/SlideRenderer';
import { parseXml } from '../../../src/parser/XmlParser';

function colorFixture(color: string, fontRef = '') {
  const fixture = createEditableTextFixture(
    undefined,
    '',
    `<a:r><a:rPr sz="2400">${color}</a:rPr><a:t>Colored</a:t></a:r>`,
    { empty: true },
  );
  if (fontRef) {
    const source = fixture.slide.nodes[0].source.element!;
    source.appendChild(
      parseXml(
        `<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fontRef idx="minor">${fontRef}</a:fontRef></p:style>`,
      ).element!,
    );
  }
  return fixture;
}

describe('editable text color inheritance', () => {
  it('uses the slide color mapping when creating a new paragraph', () => {
    const { presentation, slide } = colorFixture(
      '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>',
    );
    slide.colorMapOverrideMode = 'override';
    slide.colorMapOverride = new Map([['tx1', 'accent1']]);
    presentation.slideToLayout.set(0, 'layout');
    presentation.layoutToMaster.set('layout', 'master');
    presentation.masterToTheme.set('master', 'theme');
    presentation.themes.set('theme', {
      colorScheme: new Map([
        ['dk1', '000000'],
        ['accent1', 'FF0000'],
      ]),
      majorFont: { latin: 'Calibri', ea: '', cs: '' },
      minorFont: { latin: 'Calibri', ea: '', cs: '' },
      fillStyles: [],
      bgFillStyles: [],
      lineStyles: [],
      effectStyles: [],
    });
    const element = readSlideElements(presentation, slide)[0];
    const preview = renderSlide(presentation, slide);
    const paragraph = element.paragraphs![0];
    const edit = renderEditableText(presentation, slide, {
      ...element,
      paragraphs: [paragraph, { ...paragraph, runs: [{ ...paragraph.runs[0], text: '' }] }],
    });
    try {
      expect(preview.element.querySelector('span')!.style.color).toBe('rgb(255, 0, 0)');
      expect(element.paragraphs![0].runs[0].color).toBe('#FF0000');
      expect(edit.textElement.children[1].querySelector('span')!.style.color).toBe(
        'rgb(255, 0, 0)',
      );
    } finally {
      preview.dispose();
      edit.dispose();
    }
  });

  it('retains color modifiers in the style inherited by a new line', () => {
    const { presentation, slide } = colorFixture(
      '<a:solidFill><a:srgbClr val="000000"><a:tint val="0"/></a:srgbClr></a:solidFill>',
    );
    expect(readSlideElements(presentation, slide)[0].paragraphs![0].runs[0].color).toBe('#ffffff');
  });

  it('uses the shape font reference color unless the run has an explicit color', () => {
    const { presentation, slide } = colorFixture('', '<a:srgbClr val="FF0000"/>');
    expect(readSlideElements(presentation, slide)[0].paragraphs![0].runs[0].color).toBe('#FF0000');
    const explicit = colorFixture(
      '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
      '<a:srgbClr val="FF0000"/>',
    );
    expect(
      readSlideElements(explicit.presentation, explicit.slide)[0].paragraphs![0].runs[0].color,
    ).toBe('#00FF00');
  });
});
