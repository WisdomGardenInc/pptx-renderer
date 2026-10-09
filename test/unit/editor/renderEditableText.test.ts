import { createEditableTextFixture as fixture } from '../../fixtures/editable-text';
import { describe, expect, it } from 'vitest';
import { renderSlide } from '../../../src/renderer/SlideRenderer';
import { renderEditableText } from '../../../src/editor/renderEditableText';
import {
  createEditableTextRange,
  extractEditableText,
  mapEditableTextPositionToRun,
} from '../../../src/editor/editableTextDom';

describe('renderEditableText', () => {
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
        expect(input.querySelector('[data-pptx-paragraph="0"] > br')).not.toBeNull();

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
    'matches preview layout for %s and leaves other objects/background untouched',
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
        const style = (root: HTMLElement) => [
          root.style.cssText,
          ...Array.from(root.children, (child) => (child as HTMLElement).style.cssText),
        ];
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
