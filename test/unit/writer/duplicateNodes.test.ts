import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { applyPlainText, readPlainText } from '../../../src/model/nodes/textEdit';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import { duplicateNodes } from '../../../src/writer/duplicateNodes';
import { emuToPx } from '../../../src/parser/units';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const EMU = 9525;

function shape(id: number, x: number, text: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x * EMU}" y="0"/><a:ext cx="${100 * EMU}" cy="${40 * EMU}"/></a:xfrm>` +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="A87838"/></a:solidFill></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1400" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    '</p:sp>'
  );
}

function connector(id: number, from: number, to: number): string {
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Link ${id}"/>` +
    `<p:cNvCxnSpPr><a:stCxn id="${from}" idx="3"/><a:endCxn id="${to}" idx="1"/></p:cNvCxnSpPr>` +
    '<p:nvPr/></p:nvCxnSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${100 * EMU}" y="${20 * EMU}"/><a:ext cx="${50 * EMU}" cy="0"/></a:xfrm>` +
    '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>'
  );
}

function group(id: number, x: number, children: string): string {
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Step ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${x * EMU}" y="0"/><a:ext cx="${250 * EMU}" cy="${40 * EMU}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${250 * EMU}" cy="${40 * EMU}"/></a:xfrm></p:grpSpPr>` +
    children +
    '</p:grpSp>'
  );
}

function slide(body: string): SlideData {
  const xml =
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>';
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

function written(data: SlideData): Document {
  return new DOMParser().parseFromString(serializeSlide(data).xml, 'application/xml');
}

function cNvPrIds(doc: Document): string[] {
  return Array.from(doc.getElementsByTagName('*'))
    .filter((el) => el.localName === 'cNvPr')
    .map((el) => el.getAttribute('id') ?? '');
}

function topLevelNames(doc: Document): string[] {
  const tree = Array.from(doc.getElementsByTagName('*')).find((el) => el.localName === 'spTree')!;
  return Array.from(tree.children)
    .filter((el) => el.localName !== 'nvGrpSpPr' && el.localName !== 'grpSpPr')
    .map((el) => el.getElementsByTagName('*')[1]?.getAttribute('name') ?? el.localName);
}

describe('duplicateNodes', () => {
  it('copies a shape with its styling at an offset', () => {
    const data = slide(shape(2, 10, 'Quote'));
    const original = data.nodes[0];

    const [copy] = duplicateNodes(data, [original], { dx: 300, dy: 20 });

    expect(data.nodes).toHaveLength(2);
    expect(copy.position).toEqual({ x: original.position.x + 300, y: original.position.y + 20 });
    expect(copy.size).toEqual(original.size);
    expect(copy.id).not.toBe(original.id);
    const xml = serializeSlide(data).xml;
    expect(xml.match(/val="A87838"/g)).toHaveLength(2);
    expect(xml.match(/sz="1400" b="1"/g)).toHaveLength(2);
    expect(xml).toContain(`<a:off x="${(10 + 300) * EMU}" y="${20 * EMU}"/>`);
  });

  it('renumbers every shape inside a copied group so ids stay unique on the slide', () => {
    const data = slide(group(10, 0, shape(11, 0, '1') + shape(12, 150, '2') + connector(13, 11, 12)));

    duplicateNodes(data, [data.nodes[0]], { dx: 300, dy: 0 });

    const ids = cNvPrIds(written(data));
    expect(ids).toHaveLength(9);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('points connectors inside the copy at the copied shapes, not the originals', () => {
    const data = slide(group(10, 0, shape(11, 0, '1') + shape(12, 150, '2') + connector(13, 11, 12)));

    duplicateNodes(data, [data.nodes[0]], { dx: 300, dy: 0 });

    const doc = written(data);
    const groups = Array.from(doc.getElementsByTagName('*')).filter((e) => e.localName === 'grpSp');
    const copied = groups[1];
    const copiedShapeIds = Array.from(copied.getElementsByTagName('*'))
      .filter((e) => e.localName === 'cNvPr' && e.parentElement?.localName === 'nvSpPr')
      .map((e) => e.getAttribute('id'));
    const start = Array.from(copied.getElementsByTagName('*')).find((e) => e.localName === 'stCxn');
    const end = Array.from(copied.getElementsByTagName('*')).find((e) => e.localName === 'endCxn');
    expect(copiedShapeIds).toContain(start?.getAttribute('id'));
    expect(copiedShapeIds).toContain(end?.getAttribute('id'));
    expect(start?.getAttribute('id')).not.toBe('11');
  });

  it('moves a copied group as a whole and leaves its child space alone', () => {
    const data = slide(group(10, 40, shape(11, 0, '1')));
    const original = data.nodes[0];

    const [copy] = duplicateNodes(data, [original], { dx: 300, dy: 0 });

    expect(copy.position.x).toBeCloseTo(emuToPx(40 * EMU) + 300);
    const xml = serializeSlide(data).xml;
    expect(xml.match(/<a:chOff x="0" y="0"\/>/g)).toHaveLength(2);
  });

  it('inserts copies right after the last original and keeps their order', () => {
    const data = slide(shape(2, 0, 'A') + shape(3, 100, 'B') + shape(4, 200, 'tail'));
    const [a, b] = data.nodes;

    duplicateNodes(data, [a, b], { dx: 0, dy: 50 });

    expect(topLevelNames(written(data))).toEqual([
      'Shape 2',
      'Shape 3',
      'Shape 2',
      'Shape 3',
      'Shape 4',
    ]);
  });

  it('lets a copy be edited independently of its original', () => {
    const data = slide(shape(2, 10, 'Quote'));
    const original = data.nodes[0] as ShapeNodeData;

    const [copy] = duplicateNodes(data, [original], { dx: 300, dy: 0 });
    applyPlainText(copy as ShapeNodeData, 'Review');

    expect(readPlainText(original)).toBe('Quote');
    const xml = serializeSlide(data).xml;
    expect(xml).toContain('<a:t>Quote</a:t>');
    expect(xml).toContain('<a:t>Review</a:t>');
  });

  it('allocates ids past ones already held by nodes not yet written into the slide', () => {
    const data = slide(shape(2, 0, 'A'));
    duplicateNodes(data, [data.nodes[0]], { dx: 0, dy: 50 });

    duplicateNodes(data, [data.nodes[0]], { dx: 0, dy: 100 });

    const ids = cNvPrIds(written(data));
    expect(new Set(ids).size).toBe(ids.length);
  });
});
