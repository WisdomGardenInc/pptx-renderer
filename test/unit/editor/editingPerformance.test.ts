import { describe, expect, it, vi } from 'vitest';
import { createManyShapesFixture } from '../../fixtures/many-shapes';
import { createGroupedTextFixture } from '../../fixtures/grouped-text';
import { readSlideElement, readSlideElements } from '../../../src/editor/readSlideElements';
import { indexSourceElements, sourceElement } from '../../../src/editor/parts';
import { renderEditableText } from '../../../src/editor/renderEditableText';
import { fitSlideText } from '../../../src/editor/fitSlideText';
import { editSlideElements } from '../../../src/writer/ElementWriter';
import { refreshSlideParts } from '../../../src/editor/refreshSlideParts';
import { parseXml } from '../../../src/parser/XmlParser';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

describe('editing without sibling projection', () => {
  it('keeps copy snapshots by default and can omit them without changing editable values', () => {
    const { presentation, slide, elements } = createManyShapesFixture(51);
    const serialize = vi.spyOn(XMLSerializer.prototype, 'serializeToString');
    try {
      const values = readSlideElements(presentation, slide, { includeSourceXml: false });
      expect(serialize).not.toHaveBeenCalled();
      expect(values).toEqual(elements.map(({ sourceXml: _source, ...element }) => element));
      expect(structuredClone(elements)).toEqual(elements);
      expect(elements.every((element) => element.sourceXml?.includes('<p:sp'))).toBe(true);
    } finally {
      serialize.mockRestore();
    }
  });

  it('reads just the target values while retaining native draw order', () => {
    const { presentation, slide, elements } = createManyShapesFixture(401);
    const serialize = vi.spyOn(XMLSerializer.prototype, 'serializeToString');
    try {
      const element = readSlideElement(presentation, slide, '2', { includeSourceXml: false });
      expect(serialize).not.toHaveBeenCalled();
      const { sourceXml: _source, ...expected } = elements.at(-1)!;
      expect(element).toEqual(expected);
      expect(element?.zOrder).toBe(400);
      expect(readSlideElement(presentation, slide, 'missing')).toBeUndefined();
    } finally {
      serialize.mockRestore();
    }
  });

  it.each([51, 401])(
    'updates a target among %s shapes without serializing or reparsing the slide',
    (count) => {
      const { presentation, slide, element } = createManyShapesFixture(count);
      const original = new XMLSerializer().serializeToString(slide.root!.element!);
      const handle = renderEditableText(presentation, slide, element);
      const input = handle.textElement;
      const serialize = vi.spyOn(XMLSerializer.prototype, 'serializeToString');
      const parse = vi.spyOn(DOMParser.prototype, 'parseFromString');
      try {
        const draft = structuredClone(element);
        draft.paragraphs![0].runs[0].text = 'Updated target';
        handle.update(draft);
        expect(serialize).not.toHaveBeenCalled();
        expect(parse).not.toHaveBeenCalled();
        expect(handle.textElement).toBe(input);
        expect(handle.extract().paragraphs![0].runs[0].text).toBe('Updated target');
      } finally {
        serialize.mockRestore();
        parse.mockRestore();
        handle.dispose();
      }
      expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(original);
    },
  );

  it('uses a refreshed context without retaining the old source or target', () => {
    const { files, presentation, slide, elements } = createManyShapesFixture(51);
    const handle = renderEditableText(presentation, slide, elements.at(-1)!);
    const input = handle.textElement;
    try {
      const next = structuredClone(elements);
      next.at(-1)!.paragraphs![0].runs[0].text = 'Saved replacement';
      const part = editSlideElements(slide, next, elements, null, new Map());
      files.slides.set(slide.slidePath, part.xml);
      const updated = refreshSlideParts(presentation, files, [0]);
      const target = readSlideElement(updated, updated.slides[0], '2', {
        includeSourceXml: false,
      })!;
      handle.update(target, { presentation: updated, slide: updated.slides[0] });
      expect(handle.textElement).toBe(input);
      expect(handle.extract().paragraphs![0].runs[0].text).toBe('Saved replacement');
      handle.update(
        { ...target, nodeId: '3' },
        { presentation: updated, slide: updated.slides[0] },
      );
      expect(handle.textElement).toBe(input);
      expect(handle.extract().nodeId).toBe('3');
    } finally {
      handle.dispose();
    }
  });

  it('fits a batch with a bounded number of whole-slide parses', async () => {
    const { presentation, slide, elements } = createManyShapesFixture(21);
    const requested = elements.map((element) => ({
      ...element,
      autoFit: 'resize' as const,
      paragraphs: [{ runs: [{ text: 'Batch fitted text', fontSize: 24 }] }],
    }));
    const parse = vi.spyOn(DOMParser.prototype, 'parseFromString');
    try {
      const fitted = await fitSlideText(presentation, slide, requested, elements);
      expect(fitted).toHaveLength(21);
      expect(parse.mock.calls.filter(([xml]) => String(xml).includes('<p:sld'))).toHaveLength(2);
      expect(document.querySelector('[data-pptx-editable]')).toBeNull();
    } finally {
      parse.mockRestore();
    }
  });
});

describe('source indexes', () => {
  it('keeps group sources and their descendants independently addressable', () => {
    const { slide } = createGroupedTextFixture([{}, { width: 900, height: 200 }]);
    const sources = indexSourceElements(slide);
    expect(sources.get('11')?.localName).toBe('grpSp');
    expect(sourceElement(slide, '2', sources)?.localName).toBe('sp');
    expect(sourceElement(slide, '2', sources)?.parentElement?.localName).toBe('grpSp');
    expect(sourceElement(slide, '3', sources)?.parentElement?.localName).toBe('spTree');
  });

  it('retains compatibility wrappers in copy snapshots and rebuilds after tree mutations', () => {
    const { presentation, slide } = createManyShapesFixture(2);
    const tree = slide.root!.child('cSld').child('spTree').element!;
    const target = sourceElement(slide, '2')!;
    const wrapper = parseXml(
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent>',
    ).element!;
    tree.replaceChild(wrapper, target);
    wrapper.firstElementChild!.appendChild(target);
    const sources = indexSourceElements(slide);
    expect(sources.get('2')).toBe(wrapper);
    expect(
      readSlideElements(presentation, slide).find((element) => element.id === '2')?.sourceXml,
    ).toContain('AlternateContent');
    wrapper.remove();
    const copy = target.cloneNode(true) as Element;
    copy.getElementsByTagNameNS(P, 'cNvPr')[0].setAttribute('id', '99');
    tree.appendChild(copy);
    expect(sourceElement(slide, '2')).toBeUndefined();
    expect(sourceElement(slide, '99')).toBe(copy);
  });
});
