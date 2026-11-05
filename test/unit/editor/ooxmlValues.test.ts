import { describe, expect, it } from 'vitest';
import {
  createOoxmlEditorFixture,
  createPlaceholderStyleFixture,
} from '../../fixtures/ooxml-editor-values';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { parseXml } from '../../../src/parser/XmlParser';

const cases = [
  ['true', true],
  ['1', true],
  ['false', false],
  ['0', false],
] as const;

describe('OOXML editor values', () => {
  it.each(cases)(
    'reads %s consistently for text styles, locks and text box type',
    (value, enabled) => {
      const { elements } = createOoxmlEditorFixture(value);
      expect(elements[0].locked).toBe(enabled);
      expect(elements[0].type).toBe(enabled ? 'text' : 'shape');
      expect(elements[0].paragraphs![0].runs[0]).toMatchObject({ bold: enabled, italic: enabled });
    },
  );

  it('retains percentage and point units in the editable projection', () => {
    const { elements } = createOoxmlEditorFixture();
    expect(elements[0].paragraphs![1]).toMatchObject({
      spaceBefore: 2,
      spaceBeforeIsPt: false,
      spaceAfter: 12,
    });
  });

  it('edits percentage spacing and switches units without changing text styles', () => {
    const { files, presentation, slide, elements } = createOoxmlEditorFixture();
    const next = structuredClone(elements);
    const paragraph = next[0].paragraphs![1];
    paragraph.spaceBefore = 1.5;
    paragraph.spaceAfter = 0.5;
    paragraph.spaceAfterIsPt = false;
    const result = editSlideElements(slide, next, elements, null, new Map());
    const props = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .child('sp')
      .child('txBody')
      .children('p')[1]
      .child('pPr');
    expect(props.child('spcBef').child('spcPct').attr('val')).toBe('150000');
    expect(props.child('spcAft').child('spcPct').attr('val')).toBe('50000');
    expect(props.child('spcAft').child('spcPts').exists()).toBe(false);
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    expect(readSlideElements(refreshed, refreshed.slides[0])[0].paragraphs![1]).toMatchObject({
      spaceBefore: 1.5,
      spaceBeforeIsPt: false,
      spaceAfter: 0.5,
      spaceAfterIsPt: false,
      runs: elements[0].paragraphs![1].runs,
    });
  });

  it('rewrites spacing when only its unit changes, including explicit zero', () => {
    const { slide, elements } = createOoxmlEditorFixture(
      'true',
      '<a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>',
    );
    const next = structuredClone(elements);
    next[0].paragraphs![1].spaceBeforeIsPt = true;
    next[0].paragraphs![1].spaceAfterIsPt = false;
    const result = editSlideElements(slide, next, elements, null, new Map());
    const props = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .child('sp')
      .child('txBody')
      .children('p')[1]
      .child('pPr');
    expect(props.child('spcBef').child('spcPts').attr('val')).toBe('0');
    expect(props.child('spcAft').child('spcPct').attr('val')).toBe('0');
  });

  it('turns off true-valued formatting and locks after editing', () => {
    const { files, presentation, slide, elements } = createOoxmlEditorFixture();
    const next = structuredClone(elements);
    next[0].locked = false;
    Object.assign(next[0].paragraphs![0].runs[0], { bold: false, italic: false });
    const result = editSlideElements(slide, next, elements, null, new Map());
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    const after = readSlideElements(refreshed, refreshed.slides[0]);
    expect(after[0].locked).toBe(false);
    expect(after[0].paragraphs![0].runs[0]).toMatchObject({ bold: false, italic: false });
  });

  it('preserves true-valued attributes and percentage spacing in a no-op save', () => {
    const { slide, elements } = createOoxmlEditorFixture();
    expect(editSlideElements(slide, elements, elements, null, new Map()).xml).toBe(
      new XMLSerializer().serializeToString(slide.root!.element!),
    );
  });

  it.each([false, true])(
    'inherits from the layout parent type and master category (default idx: %s)',
    (defaultIndex) => {
      const { elements } = createPlaceholderStyleFixture(defaultIndex);
      expect(elements[0].paragraphs![0]).toMatchObject({
        align: 'right',
        spaceBefore: 1.5,
        spaceBeforeIsPt: false,
        spaceAfter: 6,
        runs: [{ text: 'Inherited', bold: true, italic: true, fontSize: 28 }],
      });
    },
  );
});
