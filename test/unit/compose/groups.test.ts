import { describe, expect, it } from 'vitest';
import { analyzeDeck } from '../../../src/compose/analyze';
import { composeDeck } from '../../../src/compose/composeDeck';
import { detectGroups, styleFingerprint } from '../../../src/compose/groups';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import { parseXml } from '../../../src/parser/XmlParser';
import {
  open,
  slideXml,
  stepsSlide,
  templateWithSlide,
  textShape,
  xfrm,
} from '../../fixtures/compose-template';

const PAGE = { width: 1280, height: 720 };

function parse(body: string): SlideData {
  return parseSlide(parseXml(slideXml(body)), 0, new Map(), 'ppt/slides/slide1.xml');
}

function rect(id: number, box: [number, number, number, number], fill = 'A87838'): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Rect ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></p:spPr></p:sp>`
  );
}

describe('styleFingerprint', () => {
  it('ignores position, size, text and ids but not styling', () => {
    const slide = parse(
      textShape(2, 'A', [0, 0, 100, 20], 'One', 1400) +
        textShape(3, 'B', [300, 50, 140, 30], 'Two words', 1400) +
        textShape(4, 'C', [0, 0, 100, 20], 'One', 1800),
    );
    const [a, b, c] = slide.nodes.map(styleFingerprint);

    expect(a).toBe(b);
    expect(a).not.toBe(c);
  });
});

describe('detectGroups', () => {
  it('finds the steps of a process with their roles and connector', () => {
    const slide = parseSlide(parseXml(stepsSlide()), 0, new Map(), 'ppt/slides/slide1.xml');

    const groups = detectGroups(slide.nodes, PAGE);

    expect(groups).toHaveLength(1);
    const [steps] = groups;
    expect(steps.direction).toBe('x');
    expect(steps.members).toHaveLength(4);
    expect(steps.members[0]).toEqual({ e1: '10', e2: '11', e3: '12', e4: '13' });
    expect(steps.members[3]).toEqual({ e1: '40', e2: '41', e3: '42', e4: '43' });
    expect(steps.track).toEqual(['3']);
    expect(steps.spacing).toBeCloseTo(299, 0);
    expect(steps.maxMembers).toBe(4);
  });

  it('finds a vertical list', () => {
    const rows = [0, 1, 2, 3]
      .map(
        (i) =>
          textShape(10 + i * 2, `Number ${i}`, [64, 280 + i * 67, 40, 45], `0${i + 1}`, 2400) +
          textShape(11 + i * 2, `Item ${i}`, [184, 280 + i * 67, 400, 45], `Item ${i}`, 2400 - 200),
      )
      .join('');

    const groups = detectGroups(parse(rows).nodes, PAGE);

    expect(groups).toHaveLength(1);
    expect(groups[0].direction).toBe('y');
    expect(groups[0].members).toHaveLength(4);
    expect(Object.keys(groups[0].members[0])).toHaveLength(2);
  });

  it('does not treat two lookalikes that happen to line up as a group', () => {
    const slide = parse(
      textShape(2, 'Left', [64, 600, 200, 20], 'Footer left', 1000) +
        textShape(3, 'Right', [900, 600, 200, 20], 'Footer right', 1000),
    );

    expect(detectGroups(slide.nodes, PAGE)).toEqual([]);
  });

  it('leaves a grid of identical cells to be handled as a table', () => {
    const cells = [0, 1, 2]
      .flatMap((row) =>
        [0, 1, 2].map((col) =>
          textShape(
            10 + row * 3 + col,
            `Cell ${row}${col}`,
            [80 + col * 230, 300 + row * 53, 200, 23],
            `$${row}${col}`,
            1200,
          ),
        ),
      )
      .join('');

    expect(detectGroups(parse(cells).nodes, PAGE)).toEqual([]);
  });

  it('ignores a repeated decoration that carries no text', () => {
    const stripes = [0, 1, 2, 3]
      .map((i) => rect(10 + i, [64, 280 + i * 53, 1152, 43], 'FAF6EC'))
      .join('');

    expect(detectGroups(parse(stripes).nodes, PAGE)).toEqual([]);
  });

  it('rejects members spaced unevenly', () => {
    const uneven = [64, 300, 700]
      .map((x, i) => textShape(10 + i, `Tag ${i}`, [x, 300, 150, 30], `Tag ${i}`, 1400))
      .join('');

    expect(detectGroups(parse(uneven).nodes, PAGE)).toEqual([]);
  });

  it('reports how many members fit at the original spacing', () => {
    const cards = [0, 1, 2]
      .map((i) => textShape(10 + i, `Card ${i}`, [64 + i * 200, 300, 150, 30], `Card ${i}`, 1400))
      .join('');

    const [group] = detectGroups(parse(cards).nodes, PAGE);

    expect(group.members).toHaveLength(3);
    expect(group.maxMembers).toBe(6);
  });
});

describe('analyzeDeck', () => {
  it('lists each page with its elements and groups, using the ids composeDeck acts on', async () => {
    const deck = await analyzeDeck(await templateWithSlide(stepsSlide()));

    expect(deck.pages).toHaveLength(2);
    const [first] = deck.pages;
    expect(first.index).toBe(1);
    const title = first.elements.find((element) => element.id === '2')!;
    expect(title.placeholder).toBe('title');
    expect(title.text).toBe('From quote to coverage in four steps');
    expect(title.x + title.y).toBeGreaterThan(0);
    const label = first.elements.find((element) => element.id === '12')!;
    expect(label).toMatchObject({ kind: 'shape', text: 'Quote', fontSize: 20 });
    expect(first.groups).toHaveLength(1);
  });

  it('describes groups that composeDeck can fill as they are', async () => {
    const template = await templateWithSlide(stepsSlide());
    const [page] = (await analyzeDeck(template)).pages;

    const result = await composeDeck(template, [
      {
        source: page.index,
        ops: [
          {
            op: 'set_items',
            group: page.groups[0],
            items: ['Review', 'Apply', 'Activate'].map((label, i) => ({
              e2: String(i + 1),
              e3: label,
              e4: `${label} detail.`,
            })),
          },
        ],
      },
    ]);

    expect(result.problems).toEqual([]);
    const after = (await analyzeDeck(result.bytes)).pages[0];
    expect(after.groups[0].members).toHaveLength(3);
    expect(after.elements.filter((e) => e.text === 'Review')).toHaveLength(1);
    expect((await open(result.bytes)).slides).toHaveLength(1);
  });
});
