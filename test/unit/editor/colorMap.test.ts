import { describe, expect, it } from 'vitest';
import { createColorMapFixture, type ColorMapCase } from '../../fixtures/color-map';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { parseXml } from '../../../src/parser/XmlParser';

const cases: [ColorMapCase, string[]][] = [
  ['dark', ['#000000', '#EEECE1']],
  ['identity', ['#FFFFFF', '#1F497D']],
  ['layout', ['#FFFFFF', '#1F497D']],
  ['inherit-layout', ['#FFFFFF', '#1F497D']],
  ['slide', ['#1F497D', '#FFFFFF']],
];

describe('editor color mapping', () => {
  it.each(cases)(
    'resolves fills, outlines, text and bullet colors through %s mapping',
    (mapping, expected) => {
      const { elements } = createColorMapFixture(mapping);
      expect(elements.map((element) => element.fillColor)).toEqual(expected);
      expect(elements.map((element) => element.strokeColor)).toEqual(expected);
      expect(elements.map((element) => element.paragraphs![0].runs[0].color)).toEqual(
        [...expected].reverse(),
      );
      expect(elements.map((element) => element.paragraphs![0].bullet!.color)).toEqual(
        [...expected].reverse(),
      );
    },
  );

  it.each(cases)(
    'retains scheme bindings and mapped colors after an alpha-only edit with %s mapping',
    (mapping, expected) => {
      const { files, presentation, slide, elements } = createColorMapFixture(mapping);
      const original = parseXml(files.slides.get(slide.slidePath)!);
      const next = structuredClone(elements).map((element) => ({ ...element, fillAlpha: 0.5 }));
      const result = editSlideElements(slide, next, elements, null, new Map());
      const shapes = parseXml(result.xml).child('cSld').child('spTree').children('sp');
      const serializer = new XMLSerializer();
      for (const [index, shape] of shapes.entries()) {
        const color = shape.child('spPr').child('solidFill').child('schemeClr');
        expect(color.attr('val')).toBe(index === 0 ? 'bg1' : 'tx2');
        expect(color.child('alpha').attr('val')).toBe('50000');
        expect(shape.child('spPr').child('solidFill').child('srgbClr').exists()).toBe(false);
        const before = original.child('cSld').child('spTree').children('sp')[index];
        expect(serializer.serializeToString(shape.child('txBody').element!)).toBe(
          serializer.serializeToString(before.child('txBody').element!),
        );
        expect(serializer.serializeToString(shape.child('spPr').child('ln').element!)).toBe(
          serializer.serializeToString(before.child('spPr').child('ln').element!),
        );
      }
      files.slides.set(slide.slidePath, result.xml);
      const refreshed = refreshSlideParts(presentation, files, [0]);
      const after = readSlideElements(refreshed, refreshed.slides[0]);
      expect(after.map((element) => element.fillColor)).toEqual(expected);
      expect(after.map((element) => element.fillAlpha)).toEqual([0.5, 0.5]);
      expect(after.map((element) => element.paragraphs)).toEqual(
        elements.map((element) => element.paragraphs),
      );
    },
  );

  it('preserves dark-theme XML when no fields change', () => {
    const { slide, elements } = createColorMapFixture();
    const original = new XMLSerializer().serializeToString(slide.root!.element!);
    expect(editSlideElements(slide, elements, elements, null, new Map()).xml).toBe(original);
  });
});
