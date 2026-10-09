import { expect, test } from '@playwright/test';

for (const body of [
  '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>',
  '<a:bodyPr wrap="square"><a:normAutofit fontScale="80000" lnSpcReduction="20000"/></a:bodyPr>',
  '<a:bodyPr wrap="none" anchor="ctr"/>',
  '<a:bodyPr anchor="b" vertOverflow="clip" horzOverflow="clip"><a:noAutofit/></a:bodyPr>',
]) {
  test(`Enter grows the editing input without shrinking text: ${body}`, async ({ page }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(async (body) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const { renderEditableText, renderSlide, createEditableTextRange, extractEditableText } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { files, presentation, slide, element } = createEditableTextFixture(body, '', '', {
        height: 48,
      });
      const before = new XMLSerializer().serializeToString(slide.root.element);
      const preview = renderSlide(presentation, slide, { onNodeRendered() {} });
      const edit = renderEditableText(presentation, slide, element);
      const host = document.createElement('div');
      Object.assign(host.style, {
        position: 'relative',
        whiteSpace: 'pre',
        transform: 'scale(0.75)',
        transformOrigin: 'top left',
      });
      host.append(preview.element, edit.element);
      document.body.append(host);
      const settle = () =>
        new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(resolve))),
        );
      await Promise.all([preview.ready, edit.ready, document.fonts.ready]);
      await settle();
      const root = edit.textElement;
      const metrics = (input = root) => {
        const run = input.querySelector('[data-pptx-run="0"]');
        const glyph = document.createRange();
        glyph.setStart(run.firstChild, 0);
        glyph.setEnd(run.firstChild, 1);
        const rect = input.getBoundingClientRect();
        return {
          height: rect.height,
          width: rect.width,
          fontSize: getComputedStyle(run).fontSize,
          glyphHeight: glyph.getBoundingClientRect().height,
          transform: input.style.transform,
        };
      };
      const state = {
        edit,
        preview,
        host,
        root,
        files,
        presentation,
        slide,
        before,
        draft: element,
        settle,
        metrics,
        initial: metrics(),
      };
      root.addEventListener('input', () => {
        state.draft = {
          ...state.draft,
          paragraphs: extractEditableText(root, state.draft.paragraphs),
        };
      });
      root.focus();
      const range = createEditableTextRange(root, 0, 12, 0, 12);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      return state;
    }, body);
    try {
      await page.keyboard.press('Enter');
      const afterEnter = await state.evaluate(({ metrics }) => metrics());
      await page.keyboard.insertText('Second line');
      const result = await state.evaluate(async (state) => {
        const during = state.metrics();
        state.edit.update(state.draft);
        await state.edit.ready;
        await state.settle();
        const second = state.preview.element.querySelector('[data-node-id="3"]');
        const rect = second.getBoundingClientRect();
        const hit = document.elementFromPoint(rect.left + rect.width / 2, rect.top + 10);
        return {
          initial: state.initial,
          during,
          after: state.metrics(),
          focused: document.activeElement === state.root,
          sameRoot: state.edit.textElement === state.root,
          text: state.draft.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
          unchanged:
            new XMLSerializer().serializeToString(state.slide.root.element) === state.before,
          previewHeight: second.getBoundingClientRect().height,
          otherInteractive: hit?.closest('[data-node-id]')?.getAttribute('data-node-id') === '3',
        };
      });
      for (const metrics of [afterEnter, result.during, result.after]) {
        expect(metrics.height).toBeGreaterThan(result.initial.height);
        expect(metrics.width).toBeCloseTo(result.initial.width, 1);
        expect(metrics.fontSize).toBe(result.initial.fontSize);
        expect(metrics.glyphHeight).toBeCloseTo(result.initial.glyphHeight, 1);
        expect(metrics.transform).not.toContain('scale(');
      }
      expect(result.previewHeight).toBeCloseTo(36, 1);
      expect(result.text).toBe('Hello  world\nSecond line');
      expect(result.focused).toBe(true);
      expect(result.sameRoot).toBe(true);
      expect(result.unchanged).toBe(true);
      expect(result.otherInteractive).toBe(true);

      const roundtrip = await state.evaluate(async (state) => {
        const {
          editSlideElements,
          readSlideElements,
          relationshipsPart,
          refreshSlideParts,
          renderSlide,
          renderEditableText,
        } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const committed = state.edit.extract(state.draft);
        const repeatedHeight = state.edit.extract(committed).height;
        const previous = readSlideElements(state.presentation, state.slide);
        const elements = previous.map((item) =>
          item.nodeId === committed.nodeId ? committed : item,
        );
        const relsPath = relationshipsPart(state.slide.slidePath);
        const part = editSlideElements(
          state.slide,
          elements,
          previous,
          state.files.slideRels.get(relsPath) ?? null,
          new Map(),
        );
        state.files.slides.set(state.slide.slidePath, part.xml);
        state.files.slideRels.set(relsPath, part.relsXml);
        const presentation = refreshSlideParts(state.presentation, state.files, [0]);
        const slide = presentation.slides[0];
        state.edit.dispose();
        state.edit.element.remove();
        state.preview.dispose();
        state.preview.element.remove();
        state.preview = renderSlide(presentation, slide, { onNodeRendered() {} });
        state.host.append(state.preview.element);
        await state.preview.ready;
        await state.settle();
        const measure = (root) => {
          const run = root.querySelector('[data-pptx-run="0"]') ?? root.querySelector('span');
          const glyph = document.createRange();
          glyph.setStart(run.firstChild, 0);
          glyph.setEnd(run.firstChild, 1);
          return {
            height: root.getBoundingClientRect().height,
            glyphHeight: glyph.getBoundingClientRect().height,
            fontSize: getComputedStyle(run).fontSize,
            transform: root.style.transform,
          };
        };
        const previewText = state.preview.element.querySelector('[data-node-id="2"] > div');
        const preview = measure(previewText);
        const reloaded = readSlideElements(presentation, slide).find(
          (item) => item.nodeId === committed.nodeId,
        );
        state.presentation = presentation;
        state.slide = slide;
        state.draft = reloaded;
        state.edit = renderEditableText(presentation, slide, reloaded);
        state.root = state.edit.textElement;
        state.host.append(state.edit.element);
        await state.edit.ready;
        await state.settle();
        return {
          committedHeight: committed.height,
          repeatedHeight,
          reloadedHeight: reloaded.height,
          preview,
          reentered: measure(state.edit.textElement),
          otherHeight: slide.nodes.find((node) => node.id === '3').size.h,
        };
      });
      expect(roundtrip.committedHeight).toBeGreaterThan(48);
      expect(roundtrip.repeatedHeight).toBe(roundtrip.committedHeight);
      expect(roundtrip.reloadedHeight).toBeCloseTo(roundtrip.committedHeight, 2);
      expect(roundtrip.otherHeight).toBe(48);
      for (const metrics of [roundtrip.preview, roundtrip.reentered]) {
        expect(metrics.height).toBeCloseTo(roundtrip.committedHeight * 0.75, 1);
        expect(metrics.fontSize).toBe(result.initial.fontSize);
        expect(metrics.glyphHeight).toBeCloseTo(result.initial.glyphHeight, 1);
        expect(metrics.transform).not.toContain('scale(');
      }
      await state.evaluate(async (state) => {
        const { createEditableTextRange } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const root = state.edit.textElement;
        root.focus();
        const range = createEditableTextRange(root, 0, 12, 1, 11);
        document.getSelection().removeAllRanges();
        document.getSelection().addRange(range);
      });
      await page.keyboard.press('Backspace');
      const reduced = await state.evaluate(async (state) => {
        const { editSlideElements, readSlideElements, refreshSlideParts } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        let draft = state.edit.extract(state.draft);
        state.edit.update(draft);
        await state.edit.ready;
        await state.settle();
        draft = state.edit.extract(draft);
        state.edit.update(draft);
        await state.edit.ready;
        await state.settle();
        const previous = readSlideElements(state.presentation, state.slide);
        const part = editSlideElements(
          state.slide,
          previous.map((item) => (item.nodeId === draft.nodeId ? draft : item)),
          previous,
          null,
          new Map(),
        );
        state.files.slides.set(state.slide.slidePath, part.xml);
        const refreshed = refreshSlideParts(state.presentation, state.files, [0]);
        return {
          height: draft.height,
          repeatedHeight: state.edit.extract(draft).height,
          reloadedHeight: readSlideElements(refreshed, refreshed.slides[0])[0].height,
          text: draft.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
          metrics: state.metrics(state.edit.textElement),
        };
      });
      expect(reduced.height).toBeLessThan(roundtrip.committedHeight);
      expect(reduced.repeatedHeight).toBe(reduced.height);
      expect(reduced.reloadedHeight).toBeCloseTo(reduced.height, 2);
      expect(reduced.text).toBe('Hello  world');
      expect(reduced.metrics.height).toBeLessThanOrEqual(reduced.height * 0.75);
      expect(reduced.height * 0.75 - reduced.metrics.height).toBeLessThan(0.75);
      expect(reduced.metrics.fontSize).toBe(result.initial.fontSize);
      expect(reduced.metrics.glyphHeight).toBeCloseTo(result.initial.glyphHeight, 1);
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
