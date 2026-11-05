import { describe, expect, it } from 'vitest';
import { buildPresentation, materializeAllSlideNodes } from '../../../src/model/Presentation';
import type { PptxFiles } from '../../../src/parser/ZipParser';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { editGroups, restoreUngroupedGeometry } from '../../../src/writer/groupOps';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const shape = (id: number, x: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="1000000"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:sp>`;
const group = (id: number, content: string) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipH="1"><a:off x="100000" y="200000"/><a:ext cx="7000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5143500"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
const slideXml = (content: string) =>
  `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${content}</p:spTree></p:cSld></p:sld>`;

function fixture(xml: string) {
  const path = 'ppt/slides/slide1.xml';
  const files: PptxFiles = {
    presentation: `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
    presentationRels: `<Relationships xmlns="${PKG}"><Relationship Id="r1" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`,
    contentTypes: '',
    slides: new Map([[path, xml]]),
    slideRels: new Map(),
    slideLayouts: new Map(),
    slideLayoutRels: new Map(),
    slideMasters: new Map(),
    slideMasterRels: new Map(),
    themes: new Map(),
    media: new Map(),
    charts: new Map(),
    chartStyles: new Map(),
    chartColors: new Map(),
    diagramDrawings: new Map(),
  };
  const presentation = buildPresentation(files, { lazySlides: true });
  materializeAllSlideNodes(presentation);
  return { files, presentation, slide: presentation.slides[0] };
}

describe('group editing', () => {
  it('groups siblings in order and restores them on ungroup', () => {
    const xml = slideXml(shape(2, 1000000) + shape(3, 3000000));
    const grouped = editGroups(xml, 'group', ['2', '3']);
    const doc = new DOMParser().parseFromString(grouped.xml, 'application/xml');
    expect(doc.getElementsByTagNameNS(P, 'grpSp')).toHaveLength(1);
    const result = editGroups(grouped.xml, 'ungroup', grouped.ids);
    expect(result.ids).toEqual(['2', '3']);
    const { presentation, slide } = fixture(result.xml);
    expect(readSlideElements(presentation, slide).map((element) => element.nodeId)).toEqual([
      '2',
      '3',
    ]);
  });

  it('preserves nested transformed child geometry after reparenting', () => {
    const before = fixture(slideXml(group(10, group(11, shape(2, 1000000)))));
    const previous = readSlideElements(before.presentation, before.slide);
    const result = editGroups(before.files.slides.get(before.slide.slidePath)!, 'ungroup', ['11']);
    const after = fixture(result.xml);
    const projected = readSlideElements(after.presentation, after.slide);
    const restored = restoreUngroupedGeometry(projected, previous, result.ids);
    const corrected = editSlideElements(after.slide, restored, projected, null, new Map());
    const final = fixture(corrected.xml);
    const element = readSlideElements(final.presentation, final.slide).find(
      (item) => item.nodeId === '2',
    )!;
    const original = previous.find((item) => item.nodeId === '2')!;
    for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const) {
      expect(element[key]).toBeCloseTo(original[key]!, 2);
    }
    expect(element.flipH).toBe(original.flipH);
    expect(element.flipV).toBe(original.flipV);
  });
});
