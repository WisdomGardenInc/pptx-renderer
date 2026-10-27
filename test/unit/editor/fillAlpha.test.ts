import { describe, expect, it } from 'vitest';
import { createFillAlphaFixture } from '../../fixtures/fill-alpha';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { parseXml } from '../../../src/parser/XmlParser';

function save(alpha?: number | null, color?: string | null) {
  const fixture = createFillAlphaFixture();
  const next = structuredClone(fixture.elements);
  for (const element of next) {
    if (alpha !== undefined) element.fillAlpha = alpha;
    if (color !== undefined) element.fillColor = color;
  }
  const result = editSlideElements(fixture.slide, next, fixture.elements, null, new Map());
  fixture.files.slides.set(fixture.slide.slidePath, result.xml);
  const refreshed = refreshSlideParts(fixture.presentation, fixture.files, [0]);
  return {
    ...fixture,
    shapes: parseXml(result.xml).child('cSld').child('spTree').children('sp'),
    after: readSlideElements(refreshed, refreshed.slides[0]),
  };
}

describe('fill alpha editing', () => {
  it('projects resolved colors and distinguishes non-solid and theme fills', () => {
    const { elements } = createFillAlphaFixture();
    expect(elements.map(({ fillKind, fillColor }) => ({ fillKind, fillColor }))).toEqual([
      { fillKind: 'gradient', fillColor: null },
      { fillKind: 'solid', fillColor: '#254061' },
      { fillKind: 'solid', fillColor: '#800000' },
      { fillKind: 'theme', fillColor: '#4F81BD' },
      { fillKind: 'none', fillColor: null },
      { fillKind: 'pattern', fillColor: null },
    ]);
  });

  it.each([0, 0.5, 1])(
    'changes only alpha to %s without replacing fills or color modifiers',
    (alpha) => {
      const { shapes, after } = save(alpha);
      const gradient = shapes[0].child('spPr').child('gradFill');
      expect(gradient.child('lin').attr('scaled')).toBe('1');
      expect(
        gradient
          .child('gsLst')
          .children('gs')
          .map((stop) => stop.child('srgbClr').attr('val')),
      ).toEqual(['FF0000', '0000FF']);
      expect(shapes[1].child('spPr').child('solidFill').child('schemeClr').attr('val')).toBe(
        'accent1',
      );
      expect(
        shapes[1].child('spPr').child('solidFill').child('schemeClr').child('lumMod').attr('val'),
      ).toBe('50000');
      expect(
        shapes[2].child('spPr').child('solidFill').child('srgbClr').child('lumMod').attr('val'),
      ).toBe('50000');
      expect(shapes[3].child('spPr').child('noFill').exists()).toBe(false);
      expect(shapes[3].child('style').child('fillRef').attr('idx')).toBe('1');
      expect(shapes[4].child('spPr').child('noFill').exists()).toBe(true);
      expect(shapes[5].child('spPr').child('pattFill').attr('prst')).toBe('horz');
      expect(after.map((element) => element.fillAlpha)).toEqual([
        alpha,
        alpha,
        alpha,
        alpha,
        0,
        alpha,
      ]);
      expect(after.map((element) => element.fillColor)).toEqual([
        null,
        '#254061',
        '#800000',
        '#4F81BD',
        null,
        null,
      ]);
    },
  );

  it('leaves untouched fill XML byte-for-byte equivalent', () => {
    const { elements, shapes } = save();
    const serializer = new XMLSerializer();
    for (const [index, shape] of shapes.entries()) {
      const before = parseXml(elements[index].sourceXml!);
      expect(serializer.serializeToString(shape.child('spPr').element!)).toBe(
        serializer.serializeToString(before.child('spPr').element!),
      );
    }
  });

  it('still replaces the fill when a host explicitly changes its color', () => {
    const { shapes, after } = save(0.5, '#00FF00');
    for (const shape of shapes) {
      const properties = shape.child('spPr');
      expect(properties.child('gradFill').exists()).toBe(false);
      expect(properties.child('pattFill').exists()).toBe(false);
      expect(properties.child('solidFill').child('srgbClr').attr('val')).toBe('00FF00');
    }
    expect(
      after.every((element) => element.fillColor === '#00FF00' && element.fillAlpha === 0.5),
    ).toBe(true);
  });

  it.each([
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="30000"/><a:lumMod val="70000"/></a:schemeClr></a:solidFill>',
    '<a:solidFill><a:srgbClr val="22AA66"><a:shade val="50000"/></a:srgbClr></a:solidFill>',
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="30000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>',
  ])('materializes theme paint losslessly for native opacity support: %s', (themeFill) => {
    const fixture = createFillAlphaFixture();
    const theme = fixture.presentation.themes.values().next().value!;
    theme.fillStyles[0] = parseXml(
      `<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${themeFill}</root>`,
    ).allChildren()[0];
    const originalTheme = new XMLSerializer().serializeToString(theme.fillStyles[0].element!);
    const previous = readSlideElements(fixture.presentation, fixture.slide);
    const next = structuredClone(previous);
    next[3].fillAlpha = 0.5;
    const result = editSlideElements(fixture.slide, next, previous, null, new Map());
    fixture.files.slides.set(fixture.slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(fixture.presentation, fixture.files, [0]);
    const after = readSlideElements(refreshed, refreshed.slides[0]);
    expect(after[3].fillColor).toBe(previous[3].fillColor);
    expect(after[3].fillAlpha).toBe(0.5);
    const shape = parseXml(result.xml).child('cSld').child('spTree').children('sp')[3];
    const fill = shape.child('spPr').child(theme.fillStyles[0].localName);
    expect(fill.exists()).toBe(true);
    expect(new XMLSerializer().serializeToString(fill.element!)).not.toContain('val="phClr"');
    if (themeFill.includes('tint'))
      expect(new XMLSerializer().serializeToString(fill.element!)).toContain('<a:tint val="30000"');
    if (themeFill.includes('gradFill')) expect(fill.child('lin').attr('ang')).toBe('5400000');
    expect(new XMLSerializer().serializeToString(theme.fillStyles[0].element!)).toBe(originalTheme);
    expect(after[1].fillColor).toBe(previous[1].fillColor);
    expect(after[1].fillAlpha).toBe(previous[1].fillAlpha);
  });

  it('removes prior alpha transforms when setting an absolute alpha', () => {
    const fixture = createFillAlphaFixture();
    const color = fixture.slide
      .root!.child('cSld')
      .child('spTree')
      .children('sp')[2]
      .child('spPr')
      .child('solidFill')
      .child('srgbClr').element!;
    for (const [name, value] of [
      ['alpha', '60000'],
      ['alphaMod', '50000'],
      ['alphaOff', '10000'],
    ]) {
      const modifier = color.ownerDocument.createElementNS(color.namespaceURI, `a:${name}`);
      modifier.setAttribute('val', value);
      color.appendChild(modifier);
    }
    const previous = readSlideElements(fixture.presentation, fixture.slide);
    const next = structuredClone(previous);
    next[2].fillAlpha = 0.5;
    const result = editSlideElements(fixture.slide, next, previous, null, new Map());
    const updated = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .children('sp')[2]
      .child('spPr')
      .child('solidFill')
      .child('srgbClr');
    expect(updated.child('alpha').attr('val')).toBe('50000');
    expect(updated.child('alphaMod').exists()).toBe(false);
    expect(updated.child('alphaOff').exists()).toBe(false);
    expect(updated.child('lumMod').attr('val')).toBe('50000');
  });
});
