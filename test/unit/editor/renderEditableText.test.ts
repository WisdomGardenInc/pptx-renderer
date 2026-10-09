import { createEditableTextFixture as fixture } from '../../fixtures/editable-text';
import { describe, expect, it, vi } from 'vitest';
import { renderSlide } from '../../../src/renderer/SlideRenderer';
import { renderEditableText } from '../../../src/editor/renderEditableText';
import {
  createEditableTextRange,
  extractEditableText,
  mapEditableTextPositionToRun,
} from '../../../src/editor/editableTextDom';

describe('renderEditableText', () => {
  it('extracts text and native height without mutating the draft or accumulating growth', () => {
    const { presentation, slide, element } = fixture(undefined, '', '', { height: 48 });
    const before = new XMLSerializer().serializeToString(slide.root!.element!);
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const height = vi.spyOn(edit.textElement, 'offsetHeight', 'get').mockReturnValue(81);
    try {
      const source = { ...element, id: 'host-id' };
      const snapshot = structuredClone(source);
      edit.textElement.firstElementChild!.append(document.createTextNode(' typed'));
      const value = edit.extract(source);
      expect(value.id).toBe('host-id');
      expect(value.height).toBe(81);
      expect(value.paragraphs![0].runs.map((run) => run.text).join('')).toBe('Hello  world typed');
      expect(value.paragraphs![0].runs[1].bold).toBe(true);
      expect(edit.extract(value).height).toBe(81);
      expect(source).toEqual(snapshot);
      expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);

      edit.update(value);
      height.mockReturnValue(100);
      expect(edit.extract().height).toBe(100);
      expect(edit.extract().id).toBe('host-id');
      edit.element.remove();
      expect(edit.extract().height).toBe(81);
      edit.dispose();
      expect(() => edit.extract()).toThrow('disposed');
    } finally {
      height.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it('retains empty-run formatting through extraction and consecutive empty paragraphs', () => {
    const { presentation, slide, element } = fixture();
    const styled = {
      ...element.paragraphs![0].runs[1],
      text: '',
      fontSize: 32,
      fontName: 'Georgia',
      color: '#FF0000',
      italic: true,
      underline: true,
    };
    const paragraph = { ...element.paragraphs![0], runs: [styled] };
    const draft = { ...element, paragraphs: [...element.paragraphs!, paragraph, paragraph] };
    const edit = renderEditableText(presentation, slide, draft);
    try {
      const empty = edit.textElement.children[2];
      const run = empty.querySelector<HTMLElement>('[data-pptx-run="0"]');
      expect(run).not.toBeNull();
      expect(run!.style.fontSize).toBe('32pt');
      expect(run!.style.color).toBe('rgb(255, 0, 0)');
      expect(run!.style.fontFamily).toContain('Georgia');
      expect(run!.textContent).toBe('\u200B');
      expect(empty.querySelector('br')).toBeNull();
      expect(edit.extract().paragraphs![2].runs).toEqual([styled]);
      const range = createEditableTextRange(edit.textElement, 2, 0, 2, 0)!;
      range.insertNode(document.createTextNode('Typed'));
      expect(edit.extract().paragraphs![2].runs).toEqual([{ ...styled, text: 'Typed' }]);
    } finally {
      edit.dispose();
    }
  });

  it('shrinks after deleting text and keeps the fitted height across model updates', () => {
    const { presentation, slide, element } = fixture(undefined, '', '', { height: 120 });
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const height = vi.spyOn(edit.textElement, 'offsetHeight', 'get').mockReturnValue(48);
    try {
      expect(edit.extract().height).toBe(120);
      edit.textElement.querySelector('[data-pptx-run="1"]')!.textContent = '';
      const reduced = edit.extract();
      expect(reduced.height).toBe(48);
      expect(edit.textElement.style.minHeight).toBe('0px');
      expect(edit.extract(reduced).height).toBe(48);
      edit.update(reduced);
      expect(edit.extract().height).toBe(48);
      height.mockReturnValue(64);
      edit.textElement.firstElementChild!.append(document.createTextNode(' more'));
      const grown = edit.extract();
      expect(grown.height).toBe(64);
      edit.update(grown);
      height.mockReturnValue(32);
      edit.textElement.firstElementChild!.textContent = 'Short';
      expect(edit.extract().height).toBe(32);
    } finally {
      height.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it('shrinks when the host removes a model paragraph but retains unchanged authored geometry', () => {
    const { presentation, slide, element } = fixture(undefined, '', '', { height: 120 });
    const expanded = { ...element, paragraphs: [element.paragraphs![0], element.paragraphs![0]] };
    const edit = renderEditableText(presentation, slide, expanded);
    document.body.append(edit.element);
    const height = vi.spyOn(edit.textElement, 'offsetHeight', 'get').mockReturnValue(48);
    try {
      edit.update({ ...expanded });
      expect(edit.extract().height).toBe(120);
      expect(edit.textElement.style.minHeight).toBe('120px');
      edit.update(element);
      expect(edit.textElement.style.minHeight).toBe('0px');
      expect(edit.extract().height).toBe(48);
    } finally {
      height.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it('fits both axes of unwrapped input without compounding size or changing saved XML', () => {
    const { presentation, slide, element } = fixture(
      '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>',
      '',
      '',
      { height: 48 },
    );
    const before = new XMLSerializer().serializeToString(slide.root!.element!);
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const width = vi.spyOn(edit.textElement, 'offsetWidth', 'get').mockReturnValue(801);
    const height = vi.spyOn(edit.textElement, 'offsetHeight', 'get').mockReturnValue(81);
    try {
      expect(edit.textElement.style.width).toBe('max-content');
      expect(edit.textElement.style.minWidth).toBe(`${element.width}px`);
      const grown = edit.extract();
      expect(grown.width).toBe(801);
      expect(grown.height).toBe(81);
      expect(edit.extract(grown)).toEqual(grown);
      edit.update(grown);
      expect(edit.extract()).toEqual(grown);
      width.mockReturnValue(160);
      height.mockReturnValue(32);
      edit.textElement.firstElementChild!.textContent = 'Short';
      const reduced = edit.extract();
      expect(reduced.width).toBe(160);
      expect(reduced.height).toBe(32);
      expect(edit.textElement.style.minWidth).toBe('0px');
      edit.update(reduced);
      expect(edit.extract()).toEqual(reduced);
      edit.element.remove();
      width.mockReturnValue(999);
      expect(edit.extract()).toEqual(reduced);
      expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);
    } finally {
      width.mockRestore();
      height.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it.each(['wrap="square"', ''])('retains wrapping width while fitting height (%s)', (wrap) => {
    const { presentation, slide, element } = fixture(`<a:bodyPr ${wrap}><a:spAutoFit/></a:bodyPr>`);
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const width = vi.spyOn(edit.textElement, 'offsetWidth', 'get').mockReturnValue(999);
    try {
      expect(edit.textElement.style.width).not.toBe('max-content');
      expect(edit.extract().width).toBe(element.width);
      edit.textElement.firstElementChild!.textContent = 'Short';
      expect(edit.extract().width).toBe(element.width);
    } finally {
      width.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it('does not measure height growth for vertical input', () => {
    const { presentation, slide, element } = fixture(
      '<a:bodyPr vert="eaVert"><a:noAutofit/></a:bodyPr>',
    );
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const width = vi.spyOn(edit.textElement, 'offsetWidth', 'get').mockReturnValue(999);
    const height = vi.spyOn(edit.textElement, 'offsetHeight', 'get').mockReturnValue(999);
    try {
      expect(edit.extract().width).toBe(element.width);
      expect(edit.extract().height).toBe(element.height);
      edit.textElement.firstElementChild!.textContent = 'Short';
      expect(edit.extract().height).toBe(element.height);
      expect(edit.textElement.style.minHeight).not.toBe('0px');
    } finally {
      width.mockRestore();
      height.mockRestore();
      edit.dispose();
      document.body.replaceChildren();
    }
  });

  it.each([true, false])(
    'renders an initially empty input without adding a preview text overlay (textBox=%s)',
    async (textBox) => {
      const { presentation, slide, element } = fixture(undefined, '', '', { empty: true, textBox });
      const before = new XMLSerializer().serializeToString(slide.root!.element!);
      const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
      let edit: ReturnType<typeof renderEditableText> | undefined;
      try {
        expect(element.type).toBe(textBox ? 'text' : 'shape');
        expect(preview.element.querySelector('[data-node-id="2"] > div')).toBeNull();
        expect(preview.element.querySelector('[data-node-id="2"] > svg')).not.toBeNull();
        edit = renderEditableText(presentation, slide, element);
        await edit.ready;
        expect(edit.textElement.getAttribute('contenteditable')).toBe('true');
        expect(edit.textElement.querySelector('[data-pptx-paragraph="0"] > br')).not.toBeNull();
        expect(extractEditableText(edit.textElement, element.paragraphs!)[0].runs).toEqual([
          { text: '' },
        ]);
        expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);
      } finally {
        edit?.dispose();
        preview.dispose();
      }
    },
  );

  it.each([true, false])(
    'retains the input and focus after clearing, then accepts new text (textBox=%s)',
    async (textBox) => {
      const { presentation, slide, element } = fixture(undefined, '', '', { textBox });
      const before = new XMLSerializer().serializeToString(slide.root!.element!);
      const edit = renderEditableText(presentation, slide, element);
      document.body.append(edit.element);
      const input = edit.textElement;
      try {
        input.focus();
        const empty = structuredClone(element);
        empty.paragraphs![0].runs = [{ ...empty.paragraphs![0].runs[0], text: '' }];
        edit.update(empty);
        await edit.ready;
        expect(edit.textElement).toBe(input);
        expect(document.activeElement).toBe(input);
        expect(extractEditableText(input, empty.paragraphs!)).toEqual(empty.paragraphs);
        expect(
          input.querySelector('[data-pptx-paragraph="0"] [data-pptx-run="0"]')?.textContent,
        ).toBe('\u200B');

        const range = createEditableTextRange(input, 0, 0, 0, 0)!;
        range.insertNode(document.createTextNode('New text'));
        const typed = { ...empty, paragraphs: extractEditableText(input, empty.paragraphs!) };
        edit.update(typed);
        expect(edit.textElement).toBe(input);
        expect(extractEditableText(input, typed.paragraphs)).toEqual(typed.paragraphs);
        expect(input.textContent).toBe('New text');
        expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);
      } finally {
        edit.dispose();
        document.body.replaceChildren();
      }
    },
  );

  it.each([
    '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>',
    '<a:bodyPr wrap="square"><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
    '<a:bodyPr wrap="none"/>',
    '<a:bodyPr anchor="b" tIns="95250" bIns="95250" vertOverflow="clip"><a:noAutofit/></a:bodyPr>',
  ])('uses content height rather than browser scaling for horizontal input: %s', async (body) => {
    const { presentation, slide, element } = fixture(
      body,
      '<a:buChar char="◆"/><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
      '',
      { height: 48 },
    );
    const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
    const edit = renderEditableText(presentation, slide, element);
    try {
      const input = edit.textElement;
      expect(input.style.height).toBe('auto');
      expect(input.style.minHeight).toBe('48px');
      expect(input.style.width).toBe(body.includes('wrap="none"') ? 'max-content' : '100%');
      expect(input.style.whiteSpace).toBe(body.includes('wrap="none"') ? 'nowrap' : 'normal');
      expect(input.style.overflowX).toBe('visible');
      expect(input.style.overflowY).toBe('visible');
      expect(input.style.transform).not.toContain('scale(');
      const previewText = preview.element.querySelector('[data-node-id="2"] > div') as HTMLElement;
      expect(previewText.style.height).toBe('100%');
      expect(previewText.style.minHeight).toBe('');
      const fontSize = (input.querySelector('[data-pptx-run="0"]') as HTMLElement).style.fontSize;
      const next = structuredClone(element);
      next.paragraphs!.push(structuredClone(next.paragraphs![0]));
      edit.update(next);
      await edit.ready;
      expect(edit.textElement).toBe(input);
      expect(input.style.height).toBe('auto');
      expect(input.querySelectorAll('[data-pptx-paragraph]')).toHaveLength(2);
      expect((input.querySelector('[data-pptx-run="0"]') as HTMLElement).style.fontSize).toBe(
        fontSize,
      );
      expect(input.style.transform).not.toContain('scale(');
    } finally {
      edit.dispose();
      preview.dispose();
    }
  });

  it('keeps vertical input on its native fixed-height layout', () => {
    const { presentation, slide, element } = fixture(
      '<a:bodyPr vert="eaVert" vertOverflow="clip"><a:noAutofit/></a:bodyPr>',
    );
    const edit = renderEditableText(presentation, slide, element);
    try {
      expect(edit.textElement.style.height).toBe('100%');
      expect(edit.textElement.style.minHeight).toBe('');
      expect(edit.textElement.style.writingMode).toBe('vertical-rl');
      expect(edit.textElement.style.overflowY).toBe('clip');
    } finally {
      edit.dispose();
    }
  });

  it.each([
    ['default', '<a:bodyPr><a:noAutofit/></a:bodyPr>', ''],
    [
      'percentage',
      '<a:bodyPr><a:noAutofit/></a:bodyPr>',
      '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
    ],
    [
      'autofit reduction',
      '<a:bodyPr><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
      '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
    ],
    [
      'point spacing',
      '<a:bodyPr><a:normAutofit fontScale="80000"/></a:bodyPr>',
      '<a:lnSpc><a:spcPts val="1800"/></a:lnSpc>',
    ],
    [
      'vertical',
      '<a:bodyPr vert="eaVert"><a:noAutofit/></a:bodyPr>',
      '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
    ],
    ['shape autofit', '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>', ''],
    ['no wrapping', '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>', ''],
  ])(
    'preserves preview typography for %s and leaves other objects/background untouched',
    async (_name, body, pPr) => {
      const { presentation, slide, element } = fixture(body, pPr);
      const before = new XMLSerializer().serializeToString(slide.root!.element!);
      const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
      const edit = renderEditableText(presentation, slide, element);
      document.body.append(preview.element, edit.element);
      try {
        await Promise.all([preview.ready, edit.ready]);
        const source = preview.element.querySelector('[data-node-id="2"]')!
          .lastElementChild as HTMLElement;
        const style = (root: HTMLElement) => {
          const clone = root.cloneNode(true) as HTMLElement;
          if (!clone.style.writingMode) {
            // Inputs grow rather than fitting/clipping. Preview autofit may also force nowrap.
            for (const property of [
              'width',
              'min-width',
              'height',
              'min-height',
              'overflow',
              'overflow-x',
              'overflow-y',
              'white-space',
            ])
              clone.style.removeProperty(property);
          }
          return [
            clone.style.cssText,
            ...Array.from(clone.children, (child) => (child as HTMLElement).style.cssText),
          ];
        };
        // The input-specific decorations do not affect native text layout.
        for (const property of ['pointer-events', 'outline', 'cursor'])
          edit.textElement.style.removeProperty(property);
        expect(style(edit.textElement)).toEqual(style(source));
        expect(preview.element.querySelectorAll('[data-node-id]')).toHaveLength(2);
        expect(preview.element.style.backgroundImage).toContain('background.png');
        expect(edit.element.firstElementChild!.getAttribute('style')).not.toContain(
          'background.png',
        );
        expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);
        expect(extractEditableText(edit.textElement, element.paragraphs!)).toEqual(
          element.paragraphs,
        );
        expect(preview.element.querySelector('[data-pptx-paragraph]')).toBeNull();
      } finally {
        edit.dispose();
        preview.dispose();
        document.body.replaceChildren();
      }
    },
  );

  it('retains run identities through fixed line blocks, bullets, compact groups and trailing breaks', () => {
    const { presentation, slide, element } = fixture(
      undefined,
      '<a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:buChar char="◆"/>',
      '<a:br/><a:r><a:rPr sz="2000" i="1"/><a:t>Second</a:t></a:r><a:br/>',
    );
    const edit = renderEditableText(presentation, slide, element);
    try {
      expect(edit.textElement.querySelector('[data-bullet]')?.getAttribute('contenteditable')).toBe(
        'false',
      );
      expect(extractEditableText(edit.textElement, element.paragraphs!)).toEqual(
        element.paragraphs,
      );
      const range = createEditableTextRange(edit.textElement, 0, 13, 0, 19)!;
      expect(range.toString()).toBe('Second');
      expect(
        mapEditableTextPositionToRun(range.startContainer, range.startOffset, edit.textElement),
      ).toEqual({ paraIdx: 0, runIdx: 3, charOffset: 0 });
    } finally {
      edit.dispose();
    }
  });

  it('updates a stable root, preserves selection and reads browser text outside styled runs', () => {
    const { presentation, slide, element } = fixture();
    const before = new XMLSerializer().serializeToString(slide.root!.element!);
    const edit = renderEditableText(presentation, slide, element);
    document.body.append(edit.element);
    const input = edit.textElement;
    try {
      input.focus();
      const selected = createEditableTextRange(input, 0, 7, 0, 12)!;
      document.getSelection()!.removeAllRanges();
      document.getSelection()!.addRange(selected);
      const next = structuredClone(element);
      next.paragraphs![0].runs[1].italic = true;
      edit.update(next);
      expect(edit.textElement).toBe(input);
      expect(document.activeElement).toBe(input);
      expect(document.getSelection()!.toString()).toBe('world');
      input.firstElementChild!.appendChild(document.createTextNode(' typed'));
      const result = extractEditableText(input, next.paragraphs!);
      expect(result[0].runs.map((run) => run.text).join('')).toBe('Hello  world typed');
      expect(result[0].runs[1].italic).toBe(true);
      expect(new XMLSerializer().serializeToString(slide.root!.element!)).toBe(before);
      edit.dispose();
      edit.dispose();
      expect(() => edit.update(next)).toThrow('disposed');
    } finally {
      edit.dispose();
      document.body.replaceChildren();
    }
  });
});
