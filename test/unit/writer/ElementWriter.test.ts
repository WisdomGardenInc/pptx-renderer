import { describe, expect, it } from 'vitest';
import { buildPresentation, materializeAllSlideNodes } from '../../../src/model/Presentation';
import { parseXml } from '../../../src/parser/XmlParser';
import type { PptxFiles } from '../../../src/parser/ZipParser';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { relationshipsPart } from '../../../src/editor/parts';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const xml = `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="3810000" cy="952500"/></a:xfrm>
<a:prstGeom prst="rect"/><a:noFill/><a:effectLst><a:glow rad="123"/></a:effectLst></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="95250"><a:buChar char="•"/></a:pPr>
<a:r><a:rPr lang="zh-CN" sz="2400"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld><p:extLst><p:ext uri="keep"/></p:extLst></p:sld>`;

function fixture() {
  const paths = ['ppt/slides/slide9.xml', 'ppt/slides/slide2.xml'];
  const files: PptxFiles = {
    presentation: `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="r1"/><p:sldId id="257" r:id="r2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
    presentationRels: `<Relationships xmlns="${PKG}"><Relationship Id="r1" Type="${R}/slide" Target="slides/slide9.xml"/><Relationship Id="r2" Type="${R}/slide" Target="slides/slide2.xml"/></Relationships>`,
    contentTypes: '',
    slides: new Map(paths.map((path) => [path, xml])),
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

describe('framework-free element editing', () => {
  it('uses native pixels and points and preserves untouched XML through refresh', () => {
    const { files, presentation, slide } = fixture();
    const before = readSlideElements(presentation, slide);
    expect(before[0].x).toBe(100);
    expect(before[0].paragraphs?.[0].marL).toBe(10);
    expect(before[0].paragraphs?.[0].runs[0].fontSize).toBe(24);
    const next = structuredClone(before);
    next[0].x = 120;
    next[0].paragraphs![0].runs[0].text = 'Updated';
    const result = editSlideElements(slide, next, before, null, new Map());
    expect(result.xml).toContain('lang="zh-CN"');
    expect(result.xml).toContain('rad="123"');
    expect(result.xml).toContain('uri="keep"');
    files.slides.set(slide.slidePath, result.xml);
    files.slideRels.set(relationshipsPart(slide.slidePath), result.relsXml!);
    const after = refreshSlideParts(presentation, files, [0]);
    expect(after.slides[1]).toBe(presentation.slides[1]);
    expect(after.themes).toBe(presentation.themes);
    expect(after.masters).toBe(presentation.masters);
    const elements = readSlideElements(after, after.slides[0]);
    expect(elements[0].x).toBe(120);
    expect(elements[0].paragraphs?.[0].runs[0].text).toBe('Updated');
    expect(elements[0].paragraphs?.[0].marL).toBe(10);
  });

  it('adds embedded pictures with relationships and opacity', () => {
    const { presentation, slide } = fixture();
    const before = readSlideElements(presentation, slide);
    const picture = {
      id: 'host-picture',
      type: 'image' as const,
      x: 20,
      y: 30,
      width: 100,
      height: 50,
      imageKey: 'uploaded',
      opacity: 0.5,
    };
    const result = editSlideElements(
      slide,
      [...before, picture],
      before,
      null,
      new Map([['host-picture', 'ppt/media/picture.png']]),
    );
    expect(result.relsXml).toContain('../media/picture.png');
    const root = parseXml(result.xml);
    expect(
      root
        .child('cSld')
        .child('spTree')
        .child('pic')
        .child('blipFill')
        .child('blip')
        .child('alphaModFix')
        .numAttr('amt'),
    ).toBe(50000);
    const nodeId = [...result.ids].find(([, id]) => id === 'host-picture')![0];
    expect(Number(nodeId)).toBeGreaterThan(2);
  });

  it('refreshes layout mappings and rejects invalid indices before publishing a model', () => {
    const { files, presentation, slide } = fixture();
    files.slideRels.set(
      relationshipsPart(slide.slidePath),
      `<Relationships xmlns="${PKG}"><Relationship Id="layout" Type="${R}/slideLayout" Target="../slideLayouts/new.xml"/></Relationships>`,
    );
    expect(refreshSlideParts(presentation, files, [0]).slideToLayout.get(0)).toBe(
      'ppt/slideLayouts/new.xml',
    );
    expect(() => refreshSlideParts(presentation, files, [9])).toThrow('Unknown slide index');
    expect(presentation.slideToLayout.has(0)).toBe(false);
  });
});
