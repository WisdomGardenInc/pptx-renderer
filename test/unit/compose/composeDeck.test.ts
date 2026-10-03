import { describe, expect, it } from 'vitest';
import { composeDeck } from '../../../src/compose/composeDeck';
import type { SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { readPlainText } from '../../../src/model/nodes/textEdit';
import { GROUP, open, template } from '../../fixtures/compose-template';

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
