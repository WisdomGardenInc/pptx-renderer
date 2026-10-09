import { expect, test } from '@playwright/test';

for (const textBox of [true, false]) {
  test(`typing into a saved empty paragraph retains its formatting (textBox=${textBox})`, async ({
    page,
  }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(async (textBox) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const {
        editSlideElements,
        readSlideElements,
        refreshSlideParts,
        renderEditableText,
        createEditableTextRange,
      } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { presentation, slide, files, element } = createEditableTextFixture(undefined, '', '', {
        textBox,
      });
      const styled = {
        ...element.paragraphs[0].runs[1],
        text: '',
        color: '#FF0000',
        fontSize: 32,
        fontName: 'Georgia',
        fontEa: 'Noto Sans CJK TC',
        italic: true,
        underline: true,
      };
      const paragraph = { ...element.paragraphs[0], runs: [styled] };
      const draft = { ...element, paragraphs: [...element.paragraphs, paragraph, paragraph] };
      const previous = readSlideElements(presentation, slide);
      const part = editSlideElements(
        slide,
        previous.map((item) => (item.nodeId === draft.nodeId ? draft : item)),
        previous,
        null,
        new Map(),
      );
      files.slides.set(slide.slidePath, part.xml);
      const refreshed = refreshSlideParts(presentation, files, [0]);
      const reloaded = readSlideElements(refreshed, refreshed.slides[0])[0];
      const edit = renderEditableText(refreshed, refreshed.slides[0], reloaded);
      const host = document.createElement('div');
      Object.assign(host.style, {
        position: 'relative',
        transform: 'scale(0.75)',
        transformOrigin: 'top left',
      });
      host.append(edit.element);
      document.body.append(host);
      await Promise.all([edit.ready, document.fonts.ready]);
      const root = edit.textElement;
      root.focus();
      const range = createEditableTextRange(root, 2, 0, 2, 0);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      const state = { edit, root, host, draft: reloaded, styled };
      root.addEventListener('input', () => {
        state.draft = edit.extract(state.draft);
      });
      return state;
    }, textBox);
    try {
      await page.keyboard.insertText('Continued');
      const result = await state.evaluate(async (state) => {
        const run = state.draft.paragraphs[2].runs[0];
        state.edit.update(state.draft);
        await state.edit.ready;
        const span = state.root.children[2].querySelector('[data-pptx-run="0"]');
        const computed = getComputedStyle(span);
        const expected = { ...state.styled, text: 'Continued' };
        return {
          run,
          expected,
          color: computed.color,
          fontSize: computed.fontSize,
          fontFamily: computed.fontFamily,
          fontWeight: computed.fontWeight,
          fontStyle: computed.fontStyle,
          decoration: computed.textDecorationLine,
          focused: document.activeElement === state.root,
          text: span.textContent,
        };
      });
      expect(result.run).toEqual(result.expected);
      expect(result.color).toBe('rgb(255, 0, 0)');
      expect(Number.parseFloat(result.fontSize)).toBeCloseTo((32 * 96) / 72, 2);
      expect(result.fontFamily).toContain('Georgia');
      expect(result.fontWeight).toBe('700');
      expect(result.fontStyle).toBe('italic');
      expect(result.decoration).toContain('underline');
      expect(result.text).toBe('Continued');
      expect(result.focused).toBe(true);
    } finally {
      await state.evaluate(({ edit, host }) => {
        edit.dispose();
        host.remove();
      });
      await state.dispose();
    }
  });
}
