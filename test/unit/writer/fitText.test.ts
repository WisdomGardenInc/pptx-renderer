import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { applyPlainText } from '../../../src/model/nodes/textEdit';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import { fitText } from '../../../src/writer/fitText';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const EMU = 9525;
const PAGE = { slideWidth: 1280, slideHeight: 720 };
const ZERO_INSETS = 'lIns="0" tIns="0" rIns="0" bIns="0"';

interface Box {
  id: number;
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  sz?: number;
  autofit?: string;
  wrap?: string;
  listSz?: number;
}

function box(b: Box): string {
  const rPr = b.sz ? `<a:rPr sz="${b.sz}"/>` : '';
  const wrap = b.wrap ? ` wrap="${b.wrap}"` : '';
  const autofit = b.autofit ? `<a:${b.autofit}/>` : '';
  const lst = b.listSz
    ? `<a:lstStyle><a:lvl1pPr><a:defRPr sz="${b.listSz}"/></a:lvl1pPr></a:lstStyle>`
    : '<a:lstStyle/>';
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${b.id}" name="Box ${b.id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${b.x * EMU}" y="${b.y * EMU}"/><a:ext cx="${b.w * EMU}" cy="${b.h * EMU}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    `<p:txBody><a:bodyPr ${ZERO_INSETS}${wrap}>${autofit}</a:bodyPr>${lst}` +
    `<a:p><a:r>${rPr}<a:t>${b.text}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

function slide(...boxes: Box[]): SlideData {
  const xml =
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    boxes.map(box).join('') +
    '</p:spTree></p:cSld></p:sld>';
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

function shapeAt(data: SlideData, index: number): ShapeNodeData {
  return data.nodes[index] as ShapeNodeData;
}

describe('fitText', () => {
  it('leaves text that already fits alone', () => {
    const data = slide({
      id: 2,
      x: 64,
      y: 600,
      w: 300,
      h: 20,
      text: 'Short',
      sz: 1000,
      autofit: 'spAutoFit',
    });
    const shape = shapeAt(data, 0);

    expect(fitText(data, shape, PAGE)).toEqual({ action: 'fits', scale: 1 });
    expect(shape.size).toEqual({ w: 300, h: 20 });
  });

  it('widens a one-line box into free space on its right', () => {
    const data = slide({
      id: 2,
      x: 64,
      y: 633,
      w: 153,
      h: 16,
      text: 'Prepared: 29 April 2026',
      sz: 1000,
      autofit: 'spAutoFit',
    });
    const shape = shapeAt(data, 0);
    applyPlainText(shape, 'Prepared: 30 September 2026 for the Chen household');

    const result = fitText(data, shape, PAGE);

    expect(result.action).toBe('widened');
    expect(shape.size.w).toBeGreaterThan(153);
    expect(shape.position.x + shape.size.w).toBeLessThanOrEqual(1280 * 0.95);
    expect(serializeSlide(data).xml).toContain(`<a:ext cx="${Math.round(shape.size.w * EMU)}"`);
  });

  it('stops widening at a neighbour and grows an autofit box downwards instead', () => {
    const data = slide(
      { id: 2, x: 64, y: 300, w: 200, h: 20, text: 'Lead', sz: 1400, autofit: 'spAutoFit' },
      { id: 3, x: 300, y: 300, w: 200, h: 20, text: 'Neighbour', sz: 1400 },
      { id: 4, x: 64, y: 420, w: 200, h: 20, text: 'Below', sz: 1400 },
    );
    const shape = shapeAt(data, 0);
    applyPlainText(shape, 'We walk through your family needs together in one short session');

    const result = fitText(data, shape, PAGE);

    expect(result.action).toBe('grown');
    expect(shape.position.x + shape.size.w).toBeLessThanOrEqual(300);
    expect(shape.size.h).toBeGreaterThan(20);
    expect(shape.position.y + shape.size.h).toBeLessThanOrEqual(420);
  });

  it('shrinks the font of a box that may not grow, and writes the new size', () => {
    const data = slide(
      { id: 2, x: 64, y: 300, w: 220, h: 70, text: '92%', sz: 4000, autofit: 'normAutofit' },
      { id: 3, x: 300, y: 300, w: 200, h: 70, text: '3×', sz: 4000 },
    );
    const shape = shapeAt(data, 0);
    applyPlainText(shape, '$3,800,000');

    const result = fitText(data, shape, PAGE);

    expect(result.action).toBe('shrunk');
    expect(result.scale).toBeGreaterThanOrEqual(0.7);
    expect(result.scale).toBeLessThan(1);
    const sz = Number(/<a:rPr[^>]*sz="(\d+)"/.exec(serializeSlide(data).xml)?.[1]);
    expect(sz).toBeLessThan(4000);
    expect(sz % 50).toBe(0);
  });

  it('reports text that cannot fit even at the smallest scale', () => {
    const data = slide(
      { id: 2, x: 64, y: 300, w: 100, h: 20, text: 'x', sz: 2400, autofit: 'normAutofit' },
      { id: 3, x: 170, y: 300, w: 100, h: 20, text: 'blocker', sz: 2400 },
    );
    const shape = shapeAt(data, 0);
    applyPlainText(shape, 'A sentence far too long for a tiny box that may neither grow nor widen');

    expect(fitText(data, shape, { ...PAGE, minScale: 0.7 })).toEqual({
      action: 'overflows',
      scale: 0.7,
    });
  });

  it('counts CJK characters as full-width', () => {
    const latin = slide({
      id: 2,
      x: 64,
      y: 300,
      w: 140,
      h: 30,
      text: 'x',
      sz: 2000,
      wrap: 'none',
      autofit: 'normAutofit',
    });
    const cjk = slide({
      id: 2,
      x: 64,
      y: 300,
      w: 140,
      h: 30,
      text: 'x',
      sz: 2000,
      wrap: 'none',
      autofit: 'normAutofit',
    });
    applyPlainText(shapeAt(latin, 0), 'abcdefghij');
    applyPlainText(shapeAt(cjk, 0), '保险方案保障家庭');

    fitText(latin, shapeAt(latin, 0), PAGE);
    fitText(cjk, shapeAt(cjk, 0), PAGE);

    expect(shapeAt(cjk, 0).size.w).toBeGreaterThan(shapeAt(latin, 0).size.w);
  });

  it('gives a run without its own size an explicit one when shrinking', () => {
    const data = slide(
      { id: 2, x: 64, y: 300, w: 260, h: 100, text: 'Title', listSz: 3600, autofit: 'normAutofit' },
      { id: 3, x: 340, y: 300, w: 200, h: 100, text: 'blocker', sz: 1200 },
    );
    const shape = shapeAt(data, 0);
    applyPlainText(shape, 'A much longer title');

    const result = fitText(data, shape, PAGE);

    expect(result.action).toBe('shrunk');
    const sz = Number(/<a:r><a:rPr[^>]*sz="(\d+)"/.exec(serializeSlide(data).xml)?.[1]);
    expect(sz).toBeLessThan(3600);
  });
});
