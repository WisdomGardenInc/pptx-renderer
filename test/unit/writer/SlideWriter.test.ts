import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import { emuToPx } from '../../../src/parser/units';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function slideXml(body: string): string {
  return (
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>'
  );
}

function shapeXml(id: string, x = 100, y = 200, cx = 300, cy = 400, text = 'Hello'): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" sz="1800"/><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    '</p:sp>'
  );
}

function parse(xml: string): SlideData {
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

describe('serializeSlide', () => {
  it('round-trips an untouched slide without semantic change', () => {
    const src = slideXml(shapeXml('2'));
    const { xml } = serializeSlide(parse(src));

    expect(xml).toContain('<a:off x="100" y="200"/>');
    expect(xml).toContain('<a:ext cx="300" cy="400"/>');
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
    expect(xml).toContain('<a:t>Hello</a:t>');
    expect(xml).toContain('prst="rect"');
  });

  it('emits the XML declaration so the part is a valid standalone document', () => {
    const { xml } = serializeSlide(parse(slideXml(shapeXml('2'))));
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
  });

  it('reports the part path it was parsed from', () => {
    expect(serializeSlide(parse(slideXml(shapeXml('2')))).path).toBe('ppt/slides/slide1.xml');
  });

  it('writes geometry edits back as EMU', () => {
    const slide = parse(slideXml(shapeXml('2')));
    const node = slide.nodes[0];
    node.position.x = emuToPx(914400); // exactly 1 inch
    node.position.y = emuToPx(457200);
    node.size.w = emuToPx(1828800);
    node.size.h = emuToPx(914400);

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<a:off x="914400" y="457200"/>');
    expect(xml).toContain('<a:ext cx="1828800" cy="914400"/>');
  });

  it('writes rotation and flips, and omits them at their defaults', () => {
    const slide = parse(slideXml(shapeXml('2')));
    slide.nodes[0].rotation = 45;
    slide.nodes[0].flipH = true;

    let xml = serializeSlide(slide).xml;
    expect(xml).toContain('rot="2700000"'); // 45 * 60000
    expect(xml).toContain('flipH="1"');
    expect(xml).not.toContain('flipV');

    slide.nodes[0].rotation = 0;
    slide.nodes[0].flipH = false;
    xml = serializeSlide(slide).xml;
    expect(xml).not.toContain('rot=');
    expect(xml).not.toContain('flipH=');
  });

  it('creates a:xfrm for a placeholder that inherited its box', () => {
    const noXfrm =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/>' +
      '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr/><a:t>T</a:t></a:r></a:p></p:txBody></p:sp>';
    const slide = parse(slideXml(noXfrm));
    slide.nodes[0].position.x = emuToPx(914400);
    slide.nodes[0].size.w = emuToPx(914400);

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<a:xfrm>');
    expect(xml).toContain('<a:off x="914400"');
  });
});
