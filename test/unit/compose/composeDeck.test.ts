import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { composeDeck } from '../../../src/compose/composeDeck';
import type { GroupSpec } from '../../../src/compose/types';
import { buildPresentation, materializeSlideNodes } from '../../../src/model/Presentation';
import type { PresentationData } from '../../../src/model/Presentation';
import type { SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { readPlainText } from '../../../src/model/nodes/textEdit';
import { parseZip } from '../../../src/parser/ZipParser';

const SRC = resolve(__dirname, '../../../docs/example/1-chart-and-complex/source.pptx');
const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const EMU = 9525;
const STEP = 299;
const LABELS = ['Quote', 'Apply', 'Underwrite', 'Activate'];

function xfrm(x: number, y: number, w: number, h: number): string {
  return `<a:xfrm><a:off x="${Math.round(x * EMU)}" y="${Math.round(y * EMU)}"/><a:ext cx="${Math.round(w * EMU)}" cy="${Math.round(h * EMU)}"/></a:xfrm>`;
}

function textShape(
  id: number,
  name: string,
  box: [number, number, number, number],
  text: string,
  sz = 1400,
) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:r><a:rPr sz="${sz}"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

function stepsSlide(): string {
  const title =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>From quote to coverage in four steps</a:t></a:r></a:p></p:txBody></p:sp>';
  const track =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Track"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(128, 408, 1024, 1)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    '<a:solidFill><a:srgbClr val="CFC6B4"/></a:solidFill></p:spPr></p:sp>';
  const steps = LABELS.map((label, j) => {
    const base = 10 + j * 10;
    const x = 64 + j * STEP;
    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${base}" name="Stop ${j + 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrm(x + 107, 387, 43, 43)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
      '<a:solidFill><a:srgbClr val="A87838"/></a:solidFill></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
      textShape(base + 1, `Number ${j + 1}`, [x + 107, 397, 43, 23], String(j + 1), 1000) +
      textShape(base + 2, `Label ${j + 1}`, [x, 467, 256, 45], label, 2000) +
      textShape(base + 3, `Detail ${j + 1}`, [x, 533, 256, 45], `Detail for ${label}.`, 1100)
    );
  }).join('');
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    title +
    track +
    steps +
    '</p:spTree></p:cSld></p:sld>'
  );
}

const GROUP: GroupSpec = {
  direction: 'x',
  track: ['3'],
  members: LABELS.map((_, j) => {
    const base = 10 + j * 10;
    return { stop: `${base}`, number: `${base + 1}`, label: `${base + 2}`, detail: `${base + 3}` };
  }),
};

async function template(): Promise<Uint8Array> {
  const buf = readFileSync(SRC);
  const bytes = new Uint8Array(buf.length);
  bytes.set(buf);
  const zip = await JSZip.loadAsync(bytes);
  zip.file('ppt/slides/slide1.xml', stepsSlide(), { createFolders: false });
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(bytes: Uint8Array): Promise<PresentationData> {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  const pres = buildPresentation(await parseZip(copy.buffer));
  pres.slides.forEach((slide) => materializeSlideNodes(pres, slide));
  return pres;
}

function named(slide: SlideData, prefix: string) {
  return slide.nodes.filter((node) => node.name.startsWith(prefix));
}

function textOf(slide: SlideData, id: string): string {
  return readPlainText(slide.nodes.find((node) => node.id === id) as ShapeNodeData);
}

function centreX(node: { position: { x: number }; size: { w: number } }): number {
  return node.position.x + node.size.w / 2;
}

function items(labels: string[]) {
  return labels.map((label, j) => ({ number: String(j + 1), label, detail: `${label} detail.` }));
}

describe('composeDeck', () => {
  it('turns four steps into three spread across the original span', async () => {
    const source = await template();
    const before = (await open(source)).slides[0];

    const result = await composeDeck(source, [
      {
        source: 1,
        ops: [
          { op: 'set_text', element: '2', text: 'Three steps to coverage' },
          { op: 'set_items', group: GROUP, items: items(['Review', 'Apply', 'Activate']) },
        ],
      },
    ]);

    expect(result.problems).toEqual([]);
    const slide = (await open(result.bytes)).slides[0];
    const stops = named(slide, 'Stop').sort((a, b) => a.position.x - b.position.x);
    expect(stops).toHaveLength(3);
    const original = named(before, 'Stop')
      .map(centreX)
      .sort((a, b) => a - b);
    expect(centreX(stops[0])).toBeCloseTo(original[0]);
    expect(centreX(stops[2])).toBeCloseTo(original[3]);
    expect(centreX(stops[1])).toBeCloseTo((original[0] + original[3]) / 2);
    expect(named(slide, 'Track')[0].size.w).toBeCloseTo(1024);
    expect(textOf(slide, '2')).toBe('Three steps to coverage');
    expect(named(slide, 'Label').map((node) => readPlainText(node as ShapeNodeData))).toEqual([
      'Review',
      'Apply',
      'Activate',
    ]);
  });

  it('keeps an inherited placeholder where its layout puts it', async () => {
    const source = await template();
    const before = (await open(source)).slides[0].nodes.find((node) => node.id === '2')!;

    const result = await composeDeck(source, [
      { source: 1, ops: [{ op: 'set_text', element: '2', text: 'Three steps to coverage' }] },
    ]);

    const after = (await open(result.bytes)).slides[0].nodes.find((node) => node.id === '2')!;
    expect(before.position.x + before.position.y).toBeGreaterThan(0);
    expect(after.position).toEqual(before.position);
  });

  it('adds members by copying the last one and tightens spacing at the page edge', async () => {
    const result = await composeDeck(await template(), [
      {
        source: 1,
        ops: [
          {
            op: 'set_items',
            group: GROUP,
            items: items(['One', 'Two', 'Three', 'Four', 'Five', 'Six']),
          },
        ],
      },
    ]);

    const pres = await open(result.bytes);
    const slide = pres.slides[0];
    const stops = named(slide, 'Stop').sort((a, b) => a.position.x - b.position.x);
    expect(stops).toHaveLength(6);
    expect(centreX(stops[5])).toBeLessThanOrEqual(pres.width * 0.95);
    const ids = slide.nodes.map((node) => node.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(named(slide, 'Number').map((node) => readPlainText(node as ShapeNodeData))).toEqual([
      '1',
      '2',
      '3',
      '4',
      '5',
      '6',
    ]);
    expect(result.problems.some((p) => p.message.includes('tighter'))).toBe(true);
  });

  it('builds several pages from one template slide and edits each on its own', async () => {
    const result = await composeDeck(await template(), [
      { source: 2, ops: [] },
      { source: 1, ops: [{ op: 'set_text', element: '2', text: 'First copy' }] },
      { source: 1, ops: [{ op: 'set_text', element: '2', text: 'Second copy' }] },
    ]);

    const pres = await open(result.bytes);
    expect(pres.slides).toHaveLength(3);
    expect(textOf(pres.slides[1], '2')).toBe('First copy');
    expect(textOf(pres.slides[2], '2')).toBe('Second copy');
  });

  it('reports instructions it cannot carry out and still produces the deck', async () => {
    const result = await composeDeck(await template(), [
      {
        source: 1,
        ops: [
          { op: 'set_text', element: '999', text: 'nowhere' },
          { op: 'set_text', element: '10', text: 'a circle has no text body' },
          {
            op: 'set_items',
            group: GROUP,
            items: [{ label: 'Only a label' }, { label: 'Another', number: '2', detail: 'x' }],
          },
        ],
      },
    ]);

    const messages = result.problems.map((p) => p.message);
    expect(messages.some((m) => m.includes('999'))).toBe(true);
    expect(messages.some((m) => m.includes('role "number"'))).toBe(true);
    expect((await open(result.bytes)).slides).toHaveLength(1);
  });

  it('does not ask for text on decorative members that carry none', async () => {
    const result = await composeDeck(await template(), [
      { source: 1, ops: [{ op: 'set_items', group: GROUP, items: items(['A', 'B', 'C', 'D']) }] },
    ]);

    expect(result.problems.filter((p) => p.message.includes('role "stop"'))).toEqual([]);
  });

  it('records how each written text was fitted', async () => {
    const result = await composeDeck(await template(), [
      {
        source: 1,
        ops: [
          {
            op: 'set_items',
            group: GROUP,
            items: items(['A considerably longer review step', 'Apply', 'Activate']),
          },
        ],
      },
    ]);

    expect(result.texts.length).toBeGreaterThanOrEqual(9);
    expect(result.texts.every((t) => t.page === 0)).toBe(true);
    expect(result.texts.some((t) => t.action !== 'fits')).toBe(true);
  });
});
