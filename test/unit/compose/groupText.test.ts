import { describe, expect, it } from 'vitest';
import { analyzeDeck } from '../../../src/compose/analyze';
import { composeDeck } from '../../../src/compose/composeDeck';
import type { GroupNodeData } from '../../../src/model/nodes/GroupNode';
import { buildTextIndex } from '../../../src/search/TextSearch';
import { EMU, open, slideXml, templateWithSlide, textShape } from '../../fixtures/compose-template';

function groupXfrm(
  off: [number, number],
  ext: [number, number],
  chOff: [number, number],
  chExt: [number, number],
): string {
  const emu = (pair: [number, number]) => `x="${pair[0] * EMU}" y="${pair[1] * EMU}"`;
  return (
    `<a:xfrm><a:off ${emu(off)}/><a:ext cx="${ext[0] * EMU}" cy="${ext[1] * EMU}"/>` +
    `<a:chOff ${emu(chOff)}/><a:chExt cx="${chExt[0] * EMU}" cy="${chExt[1] * EMU}"/></a:xfrm>`
  );
}

function cardsSlide(): string {
  const cards = [0, 1]
    .map((i) =>
      textShape(
        20 + i,
        `Card ${i + 1}`,
        [10 + i * 120, 20, 100, 40],
        `Card text ${i + 1}`,
        1200,
      ),
    )
    .join('');
  return (
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="Cards"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr>${groupXfrm([0, 0], [2000, 1000], [0, 0], [1000, 500])}</p:grpSpPr>` +
    cards +
    '</p:grpSp>'
  );
}

function groupOf(slides: Awaited<ReturnType<typeof open>>['slides']): GroupNodeData {
  return slides[0].nodes.find((node) => node.id === '50') as GroupNodeData;
}

describe('group descendants in the page analysis', () => {
  it('lists the text inside a group with the groups it sits in', async () => {
    const analysis = await analyzeDeck(await templateWithSlide(slideXml(cardsSlide())));
    const elements = analysis.pages[0].elements;

    const card = elements.find((element) => element.id === '20');
    expect(card?.text).toBe('Card text 1');
    expect(card?.parentIds).toEqual(['50']);
    expect(card?.fontSize).toBe(12);
    expect(card?.x).toBeCloseTo(20);
    expect(card?.y).toBeCloseTo(40);
    expect(card?.w).toBeCloseTo(200);
    expect(card?.h).toBeCloseTo(80);
    expect(elements.find((element) => element.id === '50')?.parentIds).toBeUndefined();
  });
});

describe('editing text inside a group', () => {
  it('replaces and clears a group child without touching the group', async () => {
    const source = await templateWithSlide(slideXml(cardsSlide()));
    const before = groupOf((await open(source)).slides);

    const result = await composeDeck(source, [
      {
        source: 1,
        ops: [
          { op: 'set_text', element: '20', text: 'Replaced card' },
          { op: 'clear', element: '21' },
        ],
      },
    ]);

    expect(result.problems).toEqual([]);
    const composed = await open(result.bytes);
    const entries = buildTextIndex(composed);
    expect(entries.find((entry) => entry.nodeId === '20')?.text).toBe('Replaced card');
    expect(entries.find((entry) => entry.nodeId === '21')).toBeUndefined();
    const after = groupOf(composed.slides);
    expect(after.position).toEqual(before.position);
    expect(after.size).toEqual(before.size);
  });

  it('shrinks long text inside the box instead of resizing the group', async () => {
    const source = await templateWithSlide(slideXml(cardsSlide()));
    const before = groupOf((await open(source)).slides);

    const result = await composeDeck(source, [
      {
        source: 1,
        ops: [{ op: 'set_text', element: '20', text: 'A very long card text '.repeat(6) }],
      },
    ]);

    expect(result.problems).toEqual([]);
    expect(['shrunk', 'overflows']).toContain(result.texts[0].action);
    expect(groupOf((await open(result.bytes)).slides).size).toEqual(before.size);
  });

  it('reports an unknown id instead of silently keeping the template text', async () => {
    const source = await templateWithSlide(slideXml(cardsSlide()));

    const result = await composeDeck(source, [
      { source: 1, ops: [{ op: 'set_text', element: '999', text: 'nowhere' }] },
    ]);

    expect(result.problems).toEqual([
      { page: 0, message: 'element 999 is not an element of the slide' },
    ]);
  });
});
