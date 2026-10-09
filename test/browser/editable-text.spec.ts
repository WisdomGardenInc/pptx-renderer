import { expect, test } from '@playwright/test';

for (const textBox of [true, false]) {
  for (const empty of [true, false]) {
    test(`editable input accepts typing after clearing (textBox=${textBox}, empty=${empty})`, async ({
      page,
    }) => {
      await page.route('https://example.com/background.png', (route) => route.abort());
      await page.goto('/test/browser/blank.html');
      const state = await page.evaluateHandle(
        async ({ textBox, empty }) => {
          const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
          const { renderEditableText, renderSlide, extractEditableText } =
            await import('/dist/aiden0z-pptx-renderer.browser.es.js');
          const { presentation, slide, element } = createEditableTextFixture(undefined, '', '', {
            textBox,
            empty,
          });
          const before = new XMLSerializer().serializeToString(slide.root.element);
          const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
          const edit = renderEditableText(presentation, slide, element);
          const host = document.createElement('div');
          host.style.position = 'relative';
          host.style.whiteSpace = 'pre';
          host.append(preview.element, edit.element);
          document.body.append(host);
          await Promise.all([preview.ready, edit.ready]);
          const root = edit.textElement;
          const state = { edit, preview, host, root, slide, before, draft: element };
          root.addEventListener('input', () => {
            state.draft = {
              ...state.draft,
              paragraphs: extractEditableText(root, state.draft.paragraphs),
            };
          });
          root.focus();
          return state;
        },
        { textBox, empty },
      );
      try {
        if (empty) {
          expect(
            await state.evaluate(
              ({ preview }) => !!preview.element.querySelector('[data-node-id="2"] > div'),
            ),
          ).toBe(false);
          await page.keyboard.insertText('Initial input');
          expect(
            await state.evaluate(({ draft }) =>
              draft.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
            ),
          ).toBe('Initial input');
        }
        await state.evaluate(({ root }) => {
          const range = document.createRange();
          range.selectNodeContents(root);
          document.getSelection().removeAllRanges();
          document.getSelection().addRange(range);
        });
        await page.keyboard.press('Backspace');
        expect(
          await state.evaluate(({ draft }) =>
            draft.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
          ),
        ).toBe('');
        await state.evaluate(async ({ edit, draft }) => {
          edit.update(draft);
          await edit.ready;
        });
        expect(
          await state.evaluate(({ edit, root }) => ({
            sameRoot: edit.textElement === root,
            focused: document.activeElement === root,
            emptyParagraph: !!root.querySelector('[data-pptx-paragraph="0"] > br'),
          })),
        ).toEqual({ sameRoot: true, focused: true, emptyParagraph: true });
        await page.keyboard.insertText('New text');
        expect(
          await state.evaluate(({ draft }) =>
            draft.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
          ),
        ).toBe('New text');
        await state.evaluate(async ({ edit, draft }) => {
          edit.update(draft);
          await edit.ready;
        });
        expect(
          await state.evaluate(({ root, slide, before }) => ({
            text: root.textContent,
            unchanged: new XMLSerializer().serializeToString(slide.root.element) === before,
          })),
        ).toEqual({ text: 'New text', unchanged: true });
      } finally {
        await state.evaluate(({ edit, preview, host }) => {
          edit.dispose();
          preview.dispose();
          host.remove();
        });
        await state.dispose();
      }
    });
  }
}

for (const body of [
  '<a:bodyPr wrap="square"><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
  '<a:bodyPr wrap="none" vert="eaVert"><a:noAutofit/></a:bodyPr>',
]) {
  test(`editable text retains native layout and other content: ${body}`, async ({ page }) => {
    await page.route('https://example.com/background.png', (route) =>
      route.fulfill({
        contentType: 'image/png',
        body: Buffer.from(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==',
          'base64',
        ),
      }),
    );
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(async (body) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const { renderEditableText, renderSlide, createEditableTextRange, extractEditableText } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { presentation, slide, element } = createEditableTextFixture(
        body,
        '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
      );
      const before = new XMLSerializer().serializeToString(slide.root.element);
      const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
      const edit = renderEditableText(presentation, slide, element);
      const host = document.createElement('div');
      host.style.position = 'relative';
      host.style.whiteSpace = 'pre';
      host.style.width = '1600px';
      host.append(preview.element, edit.element);
      document.body.append(host);
      await Promise.all([preview.ready, edit.ready, document.fonts.ready]);
      const source = preview.element.querySelector('[data-node-id="2"]').lastElementChild;
      const metrics = (root) => {
        const p = root.firstElementChild;
        const css = getComputedStyle(p);
        const rect = p.getBoundingClientRect();
        return [
          css.lineHeight,
          css.fontSize,
          css.marginTop,
          css.marginBottom,
          getComputedStyle(root).writingMode,
          rect.width,
          rect.height,
        ];
      };
      const previewMetrics = metrics(source);
      const editMetrics = metrics(edit.textElement);
      const root = edit.textElement;
      root.focus();
      const range = createEditableTextRange(root, 0, 7, 0, 12);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      const next = structuredClone(element);
      next.paragraphs[0].runs[1].italic = true;
      edit.update(next);
      await edit.ready;
      const second = preview.element.querySelector('[data-node-id="3"]');
      const bounds = second.getBoundingClientRect();
      const hit = document.elementFromPoint(bounds.left + bounds.width / 2, bounds.top + 10);
      const actual = {
        previewMetrics,
        editMetrics,
        selection: document.getSelection().toString(),
        sameRoot: root === edit.textElement,
        focused: document.activeElement === root,
        sourceUnchanged: before === new XMLSerializer().serializeToString(slide.root.element),
        otherVisible:
          getComputedStyle(second).visibility === 'visible' &&
          hit?.closest('[data-node-id]')?.getAttribute('data-node-id') === '3',
        overlayBackground: getComputedStyle(edit.element.firstElementChild).backgroundImage,
        readText: extractEditableText(root, next.paragraphs)[0]
          .runs.map((run) => run.text)
          .join(''),
      };
      edit.dispose();
      preview.dispose();
      host.remove();
      return actual;
    }, body);
    expect(result.editMetrics).toEqual(result.previewMetrics);
    expect(result.sameRoot).toBe(true);
    expect(result.focused).toBe(true);
    expect(result.selection).toBe('world');
    expect(result.sourceUnchanged).toBe(true);
    expect(result.otherVisible).toBe(true);
    expect(result.overlayBackground).toBe('none');
    expect(result.readText).toBe('Hello  world');
  });
}
