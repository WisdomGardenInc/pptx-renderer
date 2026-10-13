import { describe, expect, it } from 'vitest';
import { buildPresentation } from '../../../src/model/Presentation';
import type { PptxFiles } from '../../../src/parser/ZipParser';
import { parseXml } from '../../../src/parser/XmlParser';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { editSlideElements } from '../../../src/writer/ElementWriter';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function fixture(
  preset = 'roundRect',
  adjustment = '<a:gd name="other" fmla="val 40000"/><a:gd name="adj" fmla="val 25000"/>',
) {
  const path = 'ppt/slides/slide1.xml';
  const xml = `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="${preset}"><a:avLst>${adjustment}</a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="112233"><a:alpha val="25000"/></a:srgbClr></a:solidFill><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="445566"><a:alpha val="60000"/></a:srgbClr></a:solidFill><a:prstDash val="dashDot"/><a:headEnd type="diamond" w="lg" len="sm"/><a:tailEnd type="triangle" w="sm" len="lg"/></a:ln><a:effectLst><a:glow rad="123"/></a:effectLst></p:spPr></p:sp></p:spTree></p:cSld><p:extLst><p:ext uri="keep"/></p:extLst></p:sld>`;
  const files: PptxFiles = {
    presentation: `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
    presentationRels: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`,
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
  const presentation = buildPresentation(files);
  const slide = presentation.slides[0];
  return { files, presentation, slide, before: readSlideElements(presentation, slide) };
}

describe('standard shape editing properties', () => {
  it('reads native radius, opacity, line caps, dash and independent endpoint sizes', () => {
    const { before } = fixture();
    expect(before[0]).toMatchObject({
      cornerRadius: 25,
      fillAlpha: 0.25,
      strokeAlpha: 0.6,
      strokeDash: 'dashdot',
      lineCap: 'rnd',
      headEnd: { type: 'diamond', w: 'lg', len: 'sm' },
      tailEnd: { type: 'triangle', w: 'sm', len: 'lg' },
    });
  });

  it('writes radius, caps and endpoint styles once while preserving other XML', () => {
    const { files, presentation, slide, before } = fixture();
    const next = [
      {
        ...before[0],
        cornerRadius: 40,
        lineCap: 'sq' as const,
        headEnd: { type: 'arrow' as const, w: 'med' as const, len: 'lg' as const },
        tailEnd: null,
        fillAlpha: 0.5,
        strokeAlpha: 0.75,
      },
    ];
    const result = editSlideElements(slide, next, before, null, new Map());
    expect(result.xml).toContain('rad="123"');
    expect(result.xml).toContain('uri="keep"');
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    expect(readSlideElements(refreshed, refreshed.slides[0])[0]).toMatchObject({
      cornerRadius: 40,
      lineCap: 'sq',
      headEnd: { type: 'arrow', w: 'med', len: 'lg' },
      fillAlpha: 0.5,
      strokeAlpha: 0.75,
    });
    const properties = parseXml(result.xml).child('cSld').child('spTree').child('sp').child('spPr');
    expect(properties.child('ln').child('tailEnd').exists()).toBe(false);
    expect(
      properties
        .child('prstGeom')
        .child('avLst')
        .children('gd')
        .find((node) => node.attr('name') === 'adj')
        ?.attr('fmla'),
    ).toBe('val 40000');
  });

  it('clamps a requested rectangle radius and restores a square rectangle at zero', () => {
    const { files, presentation, slide, before } = fixture('rect', '');
    const rounded = [{ ...before[0], cornerRadius: 1000 }];
    const result = editSlideElements(slide, rounded, before, null, new Map());
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    const radius = readSlideElements(refreshed, refreshed.slides[0]);
    expect(radius[0]).toMatchObject({ shapeType: 'roundRect', cornerRadius: 50 });
    const reset = editSlideElements(
      refreshed.slides[0],
      [{ ...radius[0], cornerRadius: 0 }],
      radius,
      null,
      new Map(),
    );
    expect(
      parseXml(reset.xml)
        .child('cSld')
        .child('spTree')
        .child('sp')
        .child('spPr')
        .child('prstGeom')
        .attr('prst'),
    ).toBe('rect');
  });

  it('applies radius to a new shape using the default rectangle preset', () => {
    const { slide, before } = fixture();
    const shape = {
      id: 'new-rounded',
      type: 'shape' as const,
      x: 0,
      y: 0,
      width: 120,
      height: 80,
      cornerRadius: 20,
    };
    const result = editSlideElements(slide, [...before, shape], before, null, new Map());
    const geometry = parseXml(result.xml)
      .child('cSld')
      .child('spTree')
      .children('sp')
      .at(-1)!
      .child('spPr')
      .child('prstGeom');
    expect(geometry.attr('prst')).toBe('roundRect');
    expect(geometry.child('avLst').child('gd').attr('fmla')).toBe('val 25000');
  });

  it('materializes independent radius and endpoint properties in a copied snapshot', () => {
    const { files, presentation, slide, before } = fixture();
    const copy = {
      ...structuredClone(before[0]),
      id: 'shape-copy',
      nodeId: undefined,
      copyStyle: 'snapshot' as const,
      cornerRadius: 40,
      lineCap: 'flat' as const,
      headEnd: { type: 'arrow' as const, w: 'sm' as const, len: 'lg' as const },
      tailEnd: null,
    };
    const result = editSlideElements(slide, [...before, copy], before, null, new Map());
    files.slides.set(slide.slidePath, result.xml);
    const refreshed = refreshSlideParts(presentation, files, [0]);
    const after = readSlideElements(refreshed, refreshed.slides[0]);
    expect(after[0]).toMatchObject({
      cornerRadius: 25,
      lineCap: 'rnd',
      headEnd: before[0].headEnd,
    });
    expect(after[1]).toMatchObject({ cornerRadius: 40, lineCap: 'flat', headEnd: copy.headEnd });
    expect(after[1].tailEnd).toBeUndefined();
    expect(after[1].nodeId).not.toBe(after[0].nodeId);
  });

  it('retains native geometry, paint and endpoint XML when no style changes are requested', () => {
    const { slide, before } = fixture('triangle');
    const original = new XMLSerializer().serializeToString(slide.root!.element!);
    const result = editSlideElements(slide, before, before, null, new Map());
    expect(result.xml).toBe(original);
  });
});
