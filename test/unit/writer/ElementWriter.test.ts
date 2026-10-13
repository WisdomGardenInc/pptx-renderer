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

  it('preserves styled empty runs through XML save and reload', () => {
    const { files, presentation, slide } = fixture();
    const previous = readSlideElements(presentation, slide);
    const next = structuredClone(previous);
    const styled = {
      ...next[0].paragraphs![0].runs[0],
      text: '',
      color: '#FF0000',
      fontSize: 32,
      fontName: 'Georgia',
      bold: true,
      italic: true,
      underline: true,
    };
    next[0].paragraphs!.push({ ...next[0].paragraphs![0], runs: [styled] });
    const result = editSlideElements(slide, next, previous, null, new Map());
    const paragraph = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .child('sp')
      .child('txBody')
      .children('p')[1];
    expect(paragraph.child('r').child('rPr').attr('sz')).toBe('3200');
    expect(paragraph.child('r').child('rPr').child('solidFill').child('srgbClr').attr('val')).toBe(
      'FF0000',
    );
    expect(paragraph.child('r').child('t').text()).toBe('');
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    expect(readSlideElements(refreshed, refreshed.slides[0])[0].paragraphs![1].runs).toEqual([
      styled,
    ]);
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

  it('materializes an explicitly requested copy snapshot without altering source styles or extensions', () => {
    const { files, presentation, slide } = fixture();
    const before = readSlideElements(presentation, slide);
    const copy = structuredClone(before[0]);
    copy.id = 'snapshot-copy';
    copy.nodeId = undefined;
    copy.copyStyle = 'snapshot';
    copy.x += 20;
    copy.paragraphs![0].runs[0] = {
      ...copy.paragraphs![0].runs[0],
      text: 'Snapshot',
      color: '#445566',
      fontName: 'Georgia',
      fontSize: 32,
      italic: true,
    };
    copy.wrap = 'none';
    copy.vAlign = 'bottom';
    copy.bodyPadding = { l: 1, r: 2, t: 3, b: 4 };
    const result = editSlideElements(slide, [...before, copy], before, null, new Map());
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    const after = readSlideElements(refreshed, refreshed.slides[0]);
    expect(after[0].paragraphs![0].runs[0].text).toBe('Hello');
    expect(after[1].paragraphs![0].runs[0]).toMatchObject(copy.paragraphs![0].runs[0]);
    expect(after[1]).toMatchObject({
      wrap: 'none',
      vAlign: 'bottom',
      bodyPadding: copy.bodyPadding,
    });
    expect(after[1].sourceXml).toContain('rad="123"');
    expect(after[1].sourceXml).toContain('lang="zh-CN"');
    expect(after[1].nodeId).not.toBe(after[0].nodeId);
  });

  it('retains original source XML style when snapshot materialization is not requested', () => {
    const { presentation, slide } = fixture();
    const before = readSlideElements(presentation, slide);
    const copy = structuredClone(before[0]);
    copy.id = 'source-copy';
    copy.nodeId = undefined;
    copy.paragraphs![0].runs[0].text = 'Not materialized';
    const result = editSlideElements(slide, [...before, copy], before, null, new Map());
    expect(result.xml).not.toContain('Not materialized');
    expect(result.xml.match(/<a:t>Hello<\/a:t>/g)).toHaveLength(2);
  });

  it('copies images and charts with inherited relationship prefixes into valid XML', () => {
    const { files, presentation, slide } = fixture();
    const chart = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>`;
    files.slides.set(slide.slidePath, xml.replace('</p:spTree>', `${chart}</p:spTree>`));
    const chartRels = `<Relationships xmlns="${PKG}"><Relationship Id="chart" Type="${R}/chart" Target="../charts/chart1.xml"/></Relationships>`;
    files.slideRels.set(relationshipsPart(slide.slidePath), chartRels);
    const withChart = refreshSlideParts(presentation, files, [0]);
    const before = readSlideElements(withChart, withChart.slides[0]);
    const picture = {
      id: 'picture',
      type: 'image' as const,
      x: 10,
      y: 20,
      width: 100,
      height: 50,
      imageKey: 'picture',
      mediaPath: 'ppt/media/picture.png',
      opacity: 0.5,
      picCrop: { l: 0.1 },
    };
    const added = editSlideElements(
      withChart.slides[0],
      [...before, picture],
      before,
      chartRels,
      new Map([[picture.id, picture.mediaPath]]),
    );
    files.slides.set(slide.slidePath, added.xml);
    files.slideRels.set(relationshipsPart(slide.slidePath), added.relsXml!);
    const baseline = refreshSlideParts(withChart, files, [0]);
    const originals = readSlideElements(baseline, baseline.slides[0]);
    const copies = originals.map((element) => ({
      ...structuredClone(element),
      id: `copy-${element.id}`,
      nodeId: undefined,
      copyStyle: 'snapshot' as const,
    }));
    const copiedPicture = copies.find((element) => element.type === 'image')!;
    copiedPicture.opacity = 0.4;
    copiedPicture.picCrop = { l: 0.2 };
    const result = editSlideElements(
      baseline.slides[0],
      [...originals, ...copies],
      originals,
      added.relsXml,
      new Map(),
    );
    expect(
      new DOMParser()
        .parseFromString(result.xml, 'application/xml')
        .getElementsByTagName('parsererror'),
    ).toHaveLength(0);
    files.slides.set(slide.slidePath, result.xml);
    files.slideRels.set(relationshipsPart(slide.slidePath), result.relsXml!);
    const refreshed = refreshSlideParts(baseline, files, [0]);
    const after = readSlideElements(refreshed, refreshed.slides[0]);
    expect(after).toHaveLength(originals.length * 2);
    expect(after.filter((element) => element.type === 'chart')).toHaveLength(2);
    expect(after.at(-1)).toMatchObject({
      opacity: 0.4,
      picCrop: { l: 0.2, r: 0, t: 0, b: 0 },
      mediaPath: picture.mediaPath,
    });
    expect(after.find((element) => element.nodeId === originals.at(-1)?.nodeId)?.opacity).toBe(0.5);
    const destination = baseline.slides[1];
    const destinationBefore = readSlideElements(baseline, destination);
    const across = editSlideElements(
      destination,
      [...destinationBefore, ...copies],
      destinationBefore,
      null,
      new Map(),
    );
    expect(
      new DOMParser()
        .parseFromString(across.xml, 'application/xml')
        .getElementsByTagName('parsererror'),
    ).toHaveLength(0);
    expect(across.relsXml).toContain('../charts/chart1.xml');
    expect(across.relsXml).toContain('../media/picture.png');
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
