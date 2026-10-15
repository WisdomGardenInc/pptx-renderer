import { expect, test } from '@playwright/test';

for (const grouped of [false, true]) {
  for (const wrap of ['square', 'none']) {
    test(`shrinks overflowing input before committing (grouped=${grouped}, wrap=${wrap})`, async ({
      page,
    }) => {
      await page.route('https://example.com/background.png', (route) => route.abort());
      await page.goto('/test/browser/blank.html');
      const state = await page.evaluateHandle(
        async ({ grouped, wrap }) => {
          const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
          const { createGroupedTextFixture } = await import('/test/fixtures/grouped-text.ts');
          const { renderEditableText, createEditableTextRange } = await import('/src/index.ts');
          const fixture = grouped
            ? createGroupedTextFixture([{}, {}])
            : createEditableTextFixture(
                `<a:bodyPr wrap="${wrap}"><a:normAutofit/></a:bodyPr>`,
                '',
                '',
                { height: 48 },
              );
          const element = { ...fixture.element, autoFit: 'shrink', wrap, height: 48 };
          const edit = renderEditableText(fixture.presentation, fixture.slide, element, {
            editableTextAutoFit: true,
          });
          const host = document.createElement('div');
          Object.assign(host.style, {
            position: 'relative',
            whiteSpace: 'pre',
            transform: 'scale(0.75)',
            transformOrigin: 'top left',
          });
          host.append(edit.element);
          document.body.append(host);
          await Promise.all([edit.ready, document.fonts.ready]);
          const root = edit.textElement;
          root.focus();
          const range = createEditableTextRange(root, 0, 12, 0, 12);
          document.getSelection().removeAllRanges();
          document.getSelection().addRange(range);
          return { ...fixture, element, edit, host, root };
        },
        { grouped, wrap },
      );
      try {
        await page.keyboard.insertText(' 文字自动缩小 Auto fit text '.repeat(16));
        // Finish with a visible glyph: native trailing spaces reserve room for the caret.
        await page.keyboard.type('.');
        await expect
          .poll(() => state.evaluate(({ root }) => Number(root.dataset.pptxFitScale ?? 1)))
          .toBeLessThan(1);
        const live = await state.evaluate((state) => ({
          initialWidth: state.element.width,
          size: state.edit.extract(),
          sameRoot: state.edit.textElement === state.root,
          focused: document.activeElement === state.root,
          scale: Number(state.root.dataset.pptxFitScale),
        }));
        expect(live.sameRoot).toBe(true);
        expect(live.focused).toBe(true);
        expect(live.size.width).toBe(live.initialWidth);
        expect(live.size.height).toBe(48);
        const parity = await state.evaluate(async (state) => {
          const { editSlideElements, readSlideElements, refreshSlideParts, renderSlide } =
            await import('/src/index.ts');
          const draft = state.edit.extract();
          const previous = readSlideElements(state.presentation, state.slide);
          const part = editSlideElements(
            state.slide,
            previous.map((item) => (item.nodeId === draft.nodeId ? draft : item)),
            previous,
            null,
            new Map(),
          );
          state.files.slides.set(state.slide.slidePath, part.xml);
          const presentation = refreshSlideParts(state.presentation, state.files, [0]);
          const preview = renderSlide(presentation, presentation.slides[0], {
            onNodeRendered() {},
          });
          state.host.append(preview.element);
          await preview.ready;
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const glyphHeight = (root) => {
            const span = root.querySelector('span');
            const range = document.createRange();
            range.setStart(span.firstChild, 0);
            range.setEnd(span.firstChild, 1);
            return range.getBoundingClientRect().height;
          };
          const positions = (root) => {
            const result = [];
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
            while (walker.nextNode()) {
              const text = walker.currentNode;
              for (let index = 0; index < text.textContent.length; index++) {
                const range = document.createRange();
                range.setStart(text, index);
                range.setEnd(text, index + 1);
                const rect = range.getBoundingClientRect();
                result.push([rect.x, rect.y, rect.width, rect.height]);
              }
            }
            return result;
          };
          const previewText = preview.element.querySelector('[data-node-id="2"] [data-pptx-text]');
          const result = {
            editing: glyphHeight(state.root),
            preview: glyphHeight(previewText),
            editingPositions: positions(state.root),
            previewPositions: positions(previewText),
          };
          preview.dispose();
          preview.element.remove();
          return result;
        });
        expect(parity.editing).toBeCloseTo(parity.preview, 1);
        expect(parity.editingPositions.length).toBe(parity.previewPositions.length);
        const positionDifferences = parity.editingPositions.flatMap((rect, index) =>
          rect.map((value, axis) => Math.abs(value - parity.previewPositions[index][axis])),
        );
        expect(Math.max(...positionDifferences)).toBeLessThanOrEqual(0.15);
        const afterUpdate = await state.evaluate(async (state) => {
          const draft = state.edit.extract();
          state.edit.update(draft);
          await state.edit.ready;
          state.root.focus();
          return {
            width: draft.width,
            height: draft.height,
            sameRoot: state.edit.textElement === state.root,
          };
        });
        expect(afterUpdate.sameRoot).toBe(true);
        expect(afterUpdate.height).toBe(48);
        await page.keyboard.press('ControlOrMeta+A');
        await page.keyboard.insertText('Short');
        await expect
          .poll(() => state.evaluate(({ root }) => Number(root.dataset.pptxFitScale ?? 1)))
          .toBe(1);
        expect(
          await state.evaluate(({ edit }) =>
            edit
              .extract()
              .paragraphs[0].runs.map((run) => run.text)
              .join(''),
          ),
        ).toBe('Short');
        const disposed = await state.evaluate(async (state) => {
          state.edit.dispose();
          const before = state.root.style.cssText;
          state.root.firstElementChild.append(document.createTextNode('After disposal'.repeat(20)));
          state.root.dispatchEvent(new InputEvent('input', { bubbles: true }));
          await new Promise((resolve) => setTimeout(resolve, 50));
          return state.root.style.cssText === before;
        });
        expect(disposed).toBe(true);
      } finally {
        await state.evaluate(({ edit, host }) => {
          edit.dispose();
          host.remove();
        });
        await state.dispose();
      }
    });
  }
}

for (const grouped of [false, true]) {
  for (const anchor of ['top', 'middle', 'bottom']) {
    test(`wrapped shrink uses available space between edits (grouped=${grouped}, anchor=${anchor})`, async ({
      page,
    }) => {
      await page.route('https://example.com/background.png', (route) => route.abort());
      await page.goto('/test/browser/blank.html');
      const state = await page.evaluateHandle(
        async ({ grouped, anchor }) => {
          const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
          const { createGroupedTextFixture } = await import('/test/fixtures/grouped-text.ts');
          const { renderEditableText, createEditableTextRange } = await import('/src/index.ts');
          const fixture = grouped
            ? createGroupedTextFixture([{}, {}])
            : createEditableTextFixture();
          const element = {
            ...fixture.element,
            width: 240,
            height: 120,
            autoFit: 'shrink',
            wrap: 'square',
            vAlign: anchor,
            paragraphs: [{ runs: [{ text: '字'.repeat(40), fontSize: 24 }] }],
          };
          const edit = renderEditableText(fixture.presentation, fixture.slide, element, {
            editableTextAutoFit: true,
          });
          const host = document.createElement('div');
          Object.assign(host.style, {
            position: 'relative',
            transform: 'scale(0.75)',
            whiteSpace: 'pre',
          });
          host.append(edit.element);
          document.body.append(host);
          await Promise.all([edit.ready, document.fonts.ready]);
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const root = edit.textElement;
          root.focus();
          const range = createEditableTextRange(root, 0, 40, 0, 40);
          document.getSelection().removeAllRanges();
          document.getSelection().addRange(range);
          return { edit, root, host, element };
        },
        { grouped, anchor },
      );
      try {
        const initial = await state.evaluate(({ root }) => {
          const scale = Number(root.dataset.pptxFitScale ?? 1);
          const width = root.style.width;
          const height = root.style.height;
          const justifyContent = root.style.justifyContent;
          const larger = Math.min(1, scale + 0.01);
          root.style.width = `${100 / larger}%`;
          root.style.height = `${100 / larger}%`;
          root.style.justifyContent = 'flex-start';
          const largerFits =
            root.scrollHeight <= root.clientHeight && root.scrollWidth <= root.clientWidth;
          root.style.width = width;
          root.style.height = height;
          root.style.justifyContent = justifyContent;
          return {
            scale,
            largerFits,
            fits: root.scrollHeight <= root.clientHeight,
            justifyContent,
          };
        });
        expect(initial.justifyContent).toBe(
          { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[anchor],
        );
        expect(initial.scale).toBeLessThan(1);
        expect(initial.fits).toBe(true);
        expect(
          initial.largerFits,
          'a larger font must overflow before shrinking is justified',
        ).toBe(false);
        await page.keyboard.insertText('.');
        await page.waitForTimeout(100);
        const after = await state.evaluate(({ root, edit }) => ({
          scale: Number(root.dataset.pptxFitScale ?? 1),
          fits: root.scrollHeight <= root.clientHeight,
          focused: document.activeElement === root,
          draft: edit.extract(),
        }));
        expect(after.scale).toBeCloseTo(initial.scale, 2);
        expect(after.fits).toBe(true);
        expect(after.focused).toBe(true);
        expect(after.draft.width).toBe(240);
        expect(after.draft.height).toBe(120);
        const acknowledged = await state.evaluate(async ({ edit, root }, draft) => {
          edit.update(draft);
          await edit.ready;
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          return {
            scale: Number(root.dataset.pptxFitScale ?? 1),
            sameRoot: edit.textElement === root,
          };
        }, after.draft);
        expect(acknowledged.scale).toBeCloseTo(after.scale, 3);
        expect(acknowledged.sameRoot).toBe(true);
        await page.keyboard.insertText('字'.repeat(60));
        await expect
          .poll(() => state.evaluate(({ root }) => Number(root.dataset.pptxFitScale ?? 1)))
          .toBeLessThan(after.scale - 0.05);
        const overflowing = await state.evaluate(({ root }) => ({
          width: root.scrollWidth > root.clientWidth + 1,
          height: root.scrollHeight > root.clientHeight,
        }));
        expect(overflowing).toEqual({ width: false, height: false });
      } finally {
        await state.evaluate(({ edit, host }) => {
          edit.dispose();
          host.remove();
        });
        await state.dispose();
      }
    });
  }
}

for (const [mode, tag] of [
  ['none', 'noAutofit'],
  ['resize', 'spAutoFit'],
] as const) {
  test(`live autofit leaves ${mode} input unscaled`, async ({ page }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(
      async ({ mode, tag }) => {
        const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
        const { renderEditableText, createEditableTextRange } = await import('/src/index.ts');
        const fixture = createEditableTextFixture(
          `<a:bodyPr wrap="none"><a:${tag}/></a:bodyPr>`,
          '',
          '',
          { height: 48 },
        );
        const edit = renderEditableText(
          fixture.presentation,
          fixture.slide,
          { ...fixture.element, autoFit: mode },
          { editableTextAutoFit: true },
        );
        document.body.append(edit.element);
        await edit.ready;
        const root = edit.textElement;
        root.focus();
        const range = createEditableTextRange(root, 0, 12, 0, 12);
        document.getSelection().removeAllRanges();
        document.getSelection().addRange(range);
        return { edit, root };
      },
      { mode, tag },
    );
    try {
      await page.keyboard.insertText('Long text '.repeat(24));
      const result = await state.evaluate(async ({ edit, root }, mode) => {
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const before = root.style.cssText;
        const draft = edit.extract();
        const measurementPreservesLayout = root.style.cssText === before;
        if (mode === 'resize') {
          edit.update(draft);
          await edit.ready;
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
        }
        return {
          width: draft.width,
          measurementPreservesLayout,
          scale: Number(root.dataset.pptxFitScale ?? 1),
          fontSize: getComputedStyle(root.querySelector('span')).fontSize,
          text: draft.paragraphs[0].runs.map((run) => run.text).join(''),
        };
      }, mode);
      expect(result.measurementPreservesLayout).toBe(true);
      if (mode === 'resize') expect(result.width).toBeGreaterThan(600);
      expect(result.scale).toBe(1);
      expect(result.fontSize).toBe('32px');
      expect(result.text).toContain('Long text');
    } finally {
      await state.evaluate(({ edit }) => {
        edit.dispose();
        edit.element.remove();
      });
      await state.dispose();
    }
  });
}
