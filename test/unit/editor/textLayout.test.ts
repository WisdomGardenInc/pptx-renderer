import { describe, expect, it } from 'vitest';
import { createEditableTextFixture } from '../../fixtures/editable-text';
import { applyTextLayout, readTextLayout, type TextAutoFit } from '../../../src/editor/textLayout';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { parseXml } from '../../../src/parser/XmlParser';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';

const modes = [
  ['none', 'noAutofit'],
  ['shrink', 'normAutofit'],
  ['resize', 'spAutoFit'],
] as const;

function bodyPr(xml: string) {
  return parseXml(xml).child('cSld').child('spTree').child('sp').child('txBody').child('bodyPr');
}

describe('editable text layout', () => {
  it('reads inherited modes and wrapping while respecting a local exclusive mode', () => {
    const { slide } = createEditableTextFixture('<a:bodyPr/>');
    const shape = slide.nodes[0] as ShapeNodeData;
    shape.textBody!.layoutBodyProperties = parseXml(
      '<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" wrap="none"><a:normAutofit fontScale="80%"/></a:bodyPr>',
    );
    expect(readTextLayout(shape.textBody)).toEqual({
      autoFit: 'shrink',
      wrap: false,
      fontScale: 0.8,
      lineSpaceReduction: null,
    });
    applyTextLayout(shape, { autoFit: 'none' });
    expect(readTextLayout(shape.textBody)).toEqual({
      autoFit: 'none',
      wrap: false,
      fontScale: null,
      lineSpaceReduction: null,
    });
    applyTextLayout(shape, { wrap: true });
    expect(readTextLayout(shape.textBody)).toEqual({
      autoFit: 'none',
      wrap: true,
      fontScale: null,
      lineSpaceReduction: null,
    });
    applyTextLayout(shape, { autoFit: null });
    expect(readTextLayout(shape.textBody).autoFit).toBe('shrink');
    expect(readTextLayout(undefined)).toEqual({
      autoFit: null,
      wrap: true,
      fontScale: null,
      lineSpaceReduction: null,
    });
  });

  it('reads inherited shrink metrics and preserves explicit zero values', () => {
    const body = {
      paragraphs: [],
      bodyProperties: parseXml(
        '<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
      ),
      layoutBodyProperties: parseXml(
        '<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:normAutofit fontScale=" 80% " lnSpcReduction="0"/></a:bodyPr>',
      ),
    };
    expect(readTextLayout(body)).toEqual({
      autoFit: 'shrink',
      wrap: true,
      fontScale: 0.8,
      lineSpaceReduction: 0,
    });
  });

  for (const [mode, tag] of modes) {
    it(`persists model ${mode} and wrap changes without discarding unrelated body properties`, () => {
      const { slide } = createEditableTextFixture(
        '<a:bodyPr anchor="ctr" lIns="123" wrap="square"><a:prstTxWarp prst="textNoShape"/><a:normAutofit fontScale="80000" lnSpcReduction="20000"/><a:extLst/></a:bodyPr>',
      );
      const shape = slide.nodes[0] as ShapeNodeData;
      applyTextLayout(shape, { autoFit: mode, wrap: false });
      const properties = bodyPr(serializeSlide(slide).xml);
      expect(properties.attr('wrap')).toBe('none');
      expect(properties.attr('anchor')).toBe('ctr');
      expect(properties.attr('lIns')).toBe('123');
      expect(properties.allChildren().map((child) => child.localName)).toEqual([
        'prstTxWarp',
        tag,
        'extLst',
      ]);
      expect(properties.child('normAutofit').attr('fontScale')).toBeUndefined();
    });
  }

  for (const [beforeMode, beforeTag] of modes) {
    for (const [mode, tag] of modes) {
      it(`roundtrips projected ${beforeMode} → ${mode} independently from wrapping`, () => {
        const { files, presentation, slide } = createEditableTextFixture(
          `<a:bodyPr wrap="none"><a:${beforeTag}${beforeMode === 'shrink' ? ' fontScale="80000"' : ''}/></a:bodyPr>`,
        );
        const previous = readSlideElements(presentation, slide);
        expect(previous[0].autoFit).toBe(beforeMode);
        const next = structuredClone(previous);
        next[0].autoFit = mode;
        next[0].wrap = 'square';
        const result = editSlideElements(slide, next, previous, null, new Map());
        const properties = bodyPr(result.xml);
        expect(properties.allChildren().map((child) => child.localName)).toEqual([tag]);
        expect(properties.attr('wrap')).toBe('square');
        expect(properties.child(tag).attr('fontScale')).toBe(
          mode === beforeMode && mode === 'shrink' ? '80000' : undefined,
        );
        files.slides.set(slide.slidePath, result.xml);
        const refreshed = refreshSlideParts(presentation, files, [0]);
        const read = readSlideElements(refreshed, refreshed.slides[0]);
        expect(read[0].autoFit).toBe(mode);
        expect(read[0].wrap).toBe('square');
      });
    }
  }

  it('preserves stored shrink metrics when only wrapping changes or a legacy host omits autoFit', () => {
    const { presentation, slide } = createEditableTextFixture(
      '<a:bodyPr><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
    );
    const previous = readSlideElements(presentation, slide);
    const next = structuredClone(previous);
    delete next[0].autoFit;
    next[0].wrap = 'none';
    const properties = bodyPr(editSlideElements(slide, next, previous, null, new Map()).xml);
    expect(properties.child('normAutofit').attr('fontScale')).toBe('80000');
    expect(properties.child('normAutofit').attr('lnSpcReduction')).toBe('20000');
  });

  it('writes the requested mode for newly created text boxes', () => {
    const { presentation, slide } = createEditableTextFixture();
    const previous = readSlideElements(presentation, slide);
    const added = {
      id: 'added',
      type: 'text' as const,
      x: 0,
      y: 0,
      width: 200,
      height: 60,
      autoFit: 'resize' as TextAutoFit,
      wrap: 'square',
      paragraphs: [{ runs: [{ text: 'Added' }] }],
    };
    const result = editSlideElements(slide, [...previous, added], previous, null, new Map());
    const properties = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .children('sp')
      .at(-1)!
      .child('txBody')
      .child('bodyPr');
    expect(properties.child('spAutoFit').exists()).toBe(true);
  });

  it('preserves authored shrink metrics when copying a text snapshot', () => {
    const { presentation, slide } = createEditableTextFixture(
      '<a:bodyPr><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
    );
    const previous = readSlideElements(presentation, slide);
    const copy = { ...previous[0], id: 'copy', nodeId: undefined, copyStyle: 'snapshot' as const };
    const part = editSlideElements(slide, [...previous, copy], previous, null, new Map());
    const fit = parseXml(part.xml)
      .child('cSld')
      .child('spTree')
      .children('sp')
      .at(-1)!
      .child('txBody')
      .child('bodyPr')
      .child('normAutofit');
    expect(fit.attr('fontScale')).toBe('80000');
    expect(fit.attr('lnSpcReduction')).toBe('20000');
  });

  it('clears font scaling while preserving line spacing and the selected mode', () => {
    const { presentation, slide } = createEditableTextFixture(
      '<a:bodyPr><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
    );
    const previous = readSlideElements(presentation, slide);
    const next = structuredClone(previous);
    next[0].fontScale = null;
    const properties = bodyPr(editSlideElements(slide, next, previous, null, new Map()).xml);
    expect(properties.child('normAutofit').exists()).toBe(true);
    expect(properties.child('normAutofit').attr('fontScale')).toBeUndefined();
    expect(properties.child('normAutofit').attr('lnSpcReduction')).toBe('20000');
    expect(previous[0].fontScale).toBe(0.8);
  });
});
