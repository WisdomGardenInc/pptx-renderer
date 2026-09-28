import { describe, expect, it } from 'vitest';
import { parseXml, SafeXmlNode } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { serializeSlide } from '../../../src/writer/SlideWriter';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function slideXml(body: string): string {
  return (
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>'
  );
}

function shapeXml(id: string, text = 'Hello'): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
    '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    '</p:sp>'
  );
}

function parse(xml: string): SlideData {
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

/** Mirrors the editor's StyleFacade: build a detached fill and reassign the ref. */
function applyFillColor(shape: ShapeNodeData, hex: string): void {
  const doc = shape.source.element!.ownerDocument;
  const sf = doc.createElementNS(A, 'a:solidFill');
  const clr = doc.createElementNS(A, 'a:srgbClr');
  clr.setAttribute('val', hex);
  sf.appendChild(clr);
  shape.fill = new SafeXmlNode(sf);
}

describe('fill sync', () => {
  it('replaces the old fill rather than appending a second one', () => {
    const slide = parse(slideXml(shapeXml('2')));
    applyFillColor(slide.nodes[0] as ShapeNodeData, '00FF00');

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<a:srgbClr val="00FF00"/>');
    expect(xml).not.toContain('val="FF0000"');
    // exactly one fill on the spPr (the a:ln keeps its own nested one)
    const spPr = xml.slice(xml.indexOf('<p:spPr>'), xml.indexOf('</p:spPr>'));
    expect(spPr.match(/<a:solidFill>/g)!.length).toBe(2); // shape fill + line fill
  });

  it('keeps a:spPr child order schema-valid (fill before a:ln)', () => {
    const slide = parse(slideXml(shapeXml('2')));
    applyFillColor(slide.nodes[0] as ShapeNodeData, '00FF00');

    const { xml } = serializeSlide(slide);
    const spPr = xml.slice(xml.indexOf('<p:spPr>'), xml.indexOf('</p:spPr>'));
    expect(spPr.indexOf('<a:xfrm>')).toBeLessThan(spPr.indexOf('<a:prstGeom'));
    expect(spPr.indexOf('<a:prstGeom')).toBeLessThan(spPr.indexOf('<a:solidFill>'));
    expect(spPr.indexOf('<a:solidFill>')).toBeLessThan(spPr.indexOf('<a:ln '));
  });

  it('places a created fill correctly on a shape that had none', () => {
    const bare =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="1"/></p:spPr></p:sp>';
    const slide = parse(slideXml(bare));
    applyFillColor(slide.nodes[0] as ShapeNodeData, '123456');

    const { xml } = serializeSlide(slide);
    const spPr = xml.slice(xml.indexOf('<p:spPr>'), xml.indexOf('</p:spPr>'));
    expect(spPr.indexOf('<a:prstGeom')).toBeLessThan(spPr.indexOf('<a:solidFill>'));
    expect(spPr.indexOf('<a:solidFill>')).toBeLessThan(spPr.indexOf('<a:ln '));
  });
});

describe('text sync', () => {
  it('writes edited run text', () => {
    const slide = parse(slideXml(shapeXml('2', 'Hello')));
    const shape = slide.nodes[0] as ShapeNodeData;
    shape.textBody!.paragraphs[0].runs[0].text = 'Goodbye';

    expect(serializeSlide(slide).xml).toContain('<a:t>Goodbye</a:t>');
  });

  it('preserves run properties across a text edit', () => {
    const slide = parse(slideXml(shapeXml('2')));
    const shape = slide.nodes[0] as ShapeNodeData;
    shape.textBody!.paragraphs[0].runs[0].text = 'New';

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('sz="1800"');
    expect(xml).toContain('b="1"');
  });

  it('rebuilds paragraphs when the editor replaces the whole body', () => {
    const slide = parse(slideXml(shapeXml('2')));
    const shape = slide.nodes[0] as ShapeNodeData;
    const style = shape.textBody!.paragraphs[0].runs[0].properties;
    // what setPlainText('A\nB') does: a fresh paragraph list reusing the first run's style
    shape.textBody!.paragraphs = [
      { level: 0, runs: [{ text: 'A', properties: style }] },
      { level: 0, runs: [{ text: 'B', properties: style }] },
    ];

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<a:t>A</a:t>');
    expect(xml).toContain('<a:t>B</a:t>');
    expect(xml).not.toContain('<a:t>Hello</a:t>');
    const txBody = xml.slice(xml.indexOf('<p:txBody>'), xml.indexOf('</p:txBody>'));
    expect(txBody.match(/<a:p>|<a:p\/>/g)!.length).toBe(2);
  });

  it('drops paragraphs removed from the model', () => {
    const twoParas = shapeXml('2').replace(
      '</p:txBody>',
      '<a:p><a:r><a:rPr/><a:t>Second</a:t></a:r></a:p></p:txBody>',
    );
    const slide = parse(slideXml(twoParas));
    const shape = slide.nodes[0] as ShapeNodeData;
    expect(shape.textBody!.paragraphs.length).toBe(2);

    shape.textBody!.paragraphs.splice(1, 1);
    const { xml } = serializeSlide(slide);
    expect(xml).not.toContain('Second');
  });
});

describe('shape tree structure', () => {
  it('removes a deleted node from the tree', () => {
    const slide = parse(slideXml(shapeXml('2', 'Keep') + shapeXml('3', 'Drop')));
    slide.nodes.splice(1, 1);

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('Keep');
    expect(xml).not.toContain('Drop');
    expect(xml).not.toContain('name="Shape 3"');
  });

  it('inserts a node created in another Document and repoints its source', () => {
    const slide = parse(slideXml(shapeXml('2')));
    // what nodeOps.createRectangle does: parse standalone XML into a fresh Document
    const created = parseSlide(parseXml(slideXml(shapeXml('99', 'Added'))), 0, new Map()).nodes[0];
    const detachedDoc = created.source.element!.ownerDocument;
    slide.nodes.push(created);

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('Added');
    expect(xml).toContain('name="Shape 99"');
    // source must now point into the slide's own Document, not the throwaway one
    expect(created.source.element!.ownerDocument).not.toBe(detachedDoc);
    expect(created.source.element!.ownerDocument).toBe(slide.root!.element!.ownerDocument);
  });

  it('a node inserted from another Document stays editable afterwards', () => {
    const slide = parse(slideXml(shapeXml('2')));
    const created = parseSlide(parseXml(slideXml(shapeXml('99', 'X'))), 0, new Map()).nodes[0];
    slide.nodes.push(created);
    serializeSlide(slide);

    // second edit, after the import — must land in the file
    created.position.x = 96; // 1 inch at 96 DPI
    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<a:off x="914400"');
  });

  it('reorders the tree to match model order', () => {
    const slide = parse(slideXml(shapeXml('2', 'First') + shapeXml('3', 'Second')));
    slide.nodes.reverse();

    const { xml } = serializeSlide(slide);
    expect(xml.indexOf('Second')).toBeLessThan(xml.indexOf('First'));
  });

  it("never removes the tree's own structural children", () => {
    const slide = parse(slideXml(shapeXml('2')));
    slide.nodes.length = 0;

    const { xml } = serializeSlide(slide);
    expect(xml).toContain('<p:nvGrpSpPr>');
    expect(xml).toContain('<p:grpSpPr/>');
    expect(xml).not.toContain('name="Shape 2"');
  });
});
