import { describe, expect, it } from 'vitest';
import { createGroupedTextFixture } from '../../fixtures/grouped-text';
import { readSlideElements } from '../../../src/editor/readSlideElements';
import { renderEditableText } from '../../../src/editor/renderEditableText';
import { sourceElement } from '../../../src/editor/parts';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { parseXml } from '../../../src/parser/XmlParser';

function save(fixture: ReturnType<typeof createGroupedTextFixture>, next = fixture.elements) {
  const result = editSlideElements(fixture.slide, next, fixture.elements, null, new Map());
  fixture.files.slides.set(fixture.slide.slidePath, result.xml);
  const presentation = refreshSlideParts(fixture.presentation, fixture.files, [0]);
  return { result, presentation, slide: presentation.slides[0] };
}

describe('text inside groups', () => {
  it('projects nested descendants in draw order with slide-space geometry', () => {
    const fixture = createGroupedTextFixture([{}, { x: 150, y: 120, width: 300, height: 200 }]);
    expect(fixture.elements.map((element) => element.nodeId)).toEqual(['10', '11', '2', '3']);
    expect(fixture.element).toMatchObject({
      x: 250,
      y: 170,
      width: 300,
      height: 150,
      type: 'text',
      parentNodeId: '11',
    });
    expect(fixture.element.paragraphs![0].runs.map((run) => run.text).join('')).toBe(
      'Hello  world',
    );
    expect(sourceElement(fixture.slide, '2')?.localName).toBe('sp');
    expect(fixture.slide.nodes).toHaveLength(2);
  });

  it('saves nested text without ungrouping or changing untouched transforms and siblings', () => {
    const fixture = createGroupedTextFixture([
      { rotation: 30, flipH: true, width: 900 },
      { rotation: 90 },
    ]);
    const serializer = new XMLSerializer();
    const beforeSibling = serializer.serializeToString(sourceElement(fixture.slide, '3')!);
    const beforeGeometry = serializer.serializeToString(
      sourceElement(fixture.slide, '2')!.getElementsByTagName('a:xfrm')[0],
    );
    const next = structuredClone(fixture.elements);
    next.find((element) => element.nodeId === '2')!.paragraphs![0].runs[0].text =
      'Edited inside group';
    const saved = save(fixture, next);
    const tree = parseXml(saved.result.xml).child('cSld').child('spTree');
    expect(tree.children('sp')).toHaveLength(1);
    expect(
      tree
        .child('grpSp')
        .child('grpSp')
        .child('sp')
        .child('txBody')
        .child('p')
        .child('r')
        .child('t')
        .text(),
    ).toBe('Edited inside group');
    expect(saved.result.xml).toContain('rad="123"');
    expect(serializer.serializeToString(sourceElement(saved.slide, '3')!)).toBe(beforeSibling);
    expect(
      serializer.serializeToString(
        sourceElement(saved.slide, '2')!.getElementsByTagName('a:xfrm')[0],
      ),
    ).toBe(beforeGeometry);
    expect(
      readSlideElements(saved.presentation, saved.slide).find((element) => element.nodeId === '2')!
        .paragraphs![0].runs[0].text,
    ).toBe('Edited inside group');
  });

  it.each([
    [{ width: 1200, height: 200 }],
    [{ rotation: 45, flipH: true, flipV: true }],
    [
      { rotation: 90, width: 900, height: 200 },
      { rotation: 90, flipH: true },
    ],
    [
      { rotation: 30, flipH: true },
      { rotation: -15, flipV: true, width: 300 },
    ],
  ])('converts edited slide-space geometry back into group space: %j', (...groups) => {
    const fixture = createGroupedTextFixture(groups);
    const next = structuredClone(fixture.elements);
    const text = next.find((element) => element.nodeId === '2')!;
    text.x += 20;
    text.y -= 10;
    text.width += 30;
    text.height += 15;
    const saved = save(fixture, next);
    const reloaded = readSlideElements(saved.presentation, saved.slide).find(
      (element) => element.nodeId === '2',
    )!;
    for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const)
      expect(reloaded[key]).toBeCloseTo(text[key]!, 3);
    expect(reloaded.flipH).toBe(text.flipH);
    expect(reloaded.flipV).toBe(text.flipV);
    expect(parseXml(saved.result.xml).child('cSld').child('spTree').children('sp')).toHaveLength(1);
  });

  it.each([90, 270])(
    'round-trips a %i-degree text shape under nonuniform group scaling',
    (rotation) => {
      const fixture = createGroupedTextFixture(
        [{ width: 900, height: 200, rotation: 30, flipH: true }],
        rotation,
      );
      const next = structuredClone(fixture.elements);
      const text = next.find((element) => element.nodeId === '2')!;
      text.x += 20;
      text.y -= 10;
      text.width += 30;
      text.height += 15;
      const saved = save(fixture, next);
      const reloaded = readSlideElements(saved.presentation, saved.slide).find(
        (element) => element.nodeId === '2',
      )!;
      for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const)
        expect(reloaded[key]).toBeCloseTo(text[key]!, 3);
    },
  );

  it('copies a whole group with unique child IDs and preserves the original group', () => {
    const fixture = createGroupedTextFixture([{}, {}]);
    const group = { ...fixture.elements[0], id: 'copied-group', x: 700 };
    const saved = save(fixture, [...fixture.elements, group]);
    const elements = readSlideElements(saved.presentation, saved.slide);
    const nativeIds = elements.map((element) => element.nodeId);
    expect(new Set(nativeIds).size).toBe(nativeIds.length);
    expect(elements.filter((element) => element.type === 'text')).toHaveLength(3);
    const copyId = [...saved.result.ids].find(([, id]) => id === 'copied-group')![0];
    expect(sourceElement(saved.slide, copyId)?.parentElement?.localName).toBe('spTree');
    expect(sourceElement(saved.slide, '2')?.parentElement?.localName).toBe('grpSp');
    expect(sourceElement(saved.slide, copyId)?.getElementsByTagName('a:t')[0].textContent).toBe(
      'Hello  ',
    );
  });

  it('deletes a descendant without moving its sibling out of the group', () => {
    const fixture = createGroupedTextFixture([{}, {}]);
    const saved = save(
      fixture,
      fixture.elements.filter((element) => element.nodeId !== '2'),
    );
    expect(sourceElement(saved.slide, '2')).toBeUndefined();
    expect(sourceElement(saved.slide, '11')?.parentElement?.localName).toBe('grpSp');
    expect(sourceElement(saved.slide, '3')?.parentElement?.localName).toBe('spTree');
  });

  it('deleting a group cannot resurrect descendants still present in the host projection', () => {
    const fixture = createGroupedTextFixture([{}, {}]);
    const saved = save(
      fixture,
      fixture.elements.filter((element) => element.nodeId !== '10'),
    );
    expect(
      readSlideElements(saved.presentation, saved.slide).map((element) => element.nodeId),
    ).toEqual(['3']);
  });

  it('creates and updates an independent text overlay for a grouped shape', async () => {
    const fixture = createGroupedTextFixture([
      { rotation: 30, flipH: true },
      { width: 300, height: 200 },
    ]);
    const before = new XMLSerializer().serializeToString(fixture.slide.root!.element!);
    const edit = renderEditableText(fixture.presentation, fixture.slide, fixture.element);
    await edit.ready;
    const input = edit.textElement;
    expect(input.getAttribute('contenteditable')).toBe('true');
    expect(input.textContent).toContain('Hello');
    const wrapper = edit.element.querySelector<HTMLElement>('[data-node-id="2"]');
    expect(parseFloat(wrapper!.style.left)).toBeCloseTo(fixture.element.x);
    expect(parseFloat(wrapper!.style.top)).toBeCloseTo(fixture.element.y);
    expect(wrapper!.style.transform).toContain('rotate(30deg)');
    const next = structuredClone(fixture.element);
    next.paragraphs![0].runs[0].text = 'Updated';
    edit.update(next);
    expect(edit.textElement).toBe(input);
    expect(edit.extract().paragraphs![0].runs[0].text).toBe('Updated');
    expect(new XMLSerializer().serializeToString(fixture.slide.root!.element!)).toBe(before);
    edit.dispose();
  });
});
