import { expect, test } from '@playwright/test';

for (const [name, body, expandsWidth, expandsHeight] of [
  [
    'unwrapped noAutofit',
    '<a:bodyPr wrap="none" anchor="ctr" lIns="190500" rIns="95250"><a:noAutofit/></a:bodyPr>',
    true,
    true,
  ],
  [
    'unwrapped normAutofit',
    '<a:bodyPr wrap="none"><a:normAutofit fontScale="80000"/></a:bodyPr>',
    true,
    true,
  ],
  ['unwrapped spAutoFit', '<a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>', true, true],
  ['unwrapped CJK', '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>', true, true],
  ['unwrapped shape', '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>', true, true],
  ['wrapped', '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>', false, true],
  ['default wrap', '<a:bodyPr><a:noAutofit/></a:bodyPr>', false, true],
  ['vertical', '<a:bodyPr wrap="none" vert="eaVert"><a:noAutofit/></a:bodyPr>', false, false],
] as const) {
  test(`editable width and height survive XML roundtrip: ${name}`, async ({ page }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(
      async ({ body, name }) => {
        const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
        const { renderEditableText, createEditableTextRange } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const fixture = createEditableTextFixture(
          body,
          '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
          '',
          { height: 48, textBox: name !== 'unwrapped shape' },
        );
        const edit = renderEditableText(fixture.presentation, fixture.slide, fixture.element);
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
        const metrics = () => {
          const run = root.querySelector('[data-pptx-run="0"]');
          const glyph = document.createRange();
          glyph.setStart(run.firstChild, 0);
          glyph.setEnd(run.firstChild, 1);
          const rect = root.getBoundingClientRect();
          return {
            width: rect.width,
            height: rect.height,
            glyphHeight: glyph.getBoundingClientRect().height,
            fontSize: getComputedStyle(run).fontSize,
          };
        };
        root.focus();
        const range = createEditableTextRange(root, 0, 12, 0, 12);
        document.getSelection().removeAllRanges();
        document.getSelection().addRange(range);
        return {
          ...fixture,
          edit,
          host,
          root,
          metrics,
          initial: metrics(),
          before: new XMLSerializer().serializeToString(fixture.slide.root.element),
        };
      },
      { body, name },
    );
    try {
      const text = name === 'unwrapped CJK' ? '文字輸入自動擴展' : 'LongText';
      await page.keyboard.insertText(text.repeat(16));
      await page.keyboard.press('Enter');
      await page.keyboard.insertText('Second line');
      const grown = await state.evaluate(async (state) => {
        const value = state.edit.extract();
        const repeated = state.edit.extract(value);
        const beforeUpdate = state.metrics();
        state.edit.update(value);
        await state.edit.ready;
        return {
          value,
          repeated,
          initial: state.initial,
          beforeUpdate,
          afterUpdate: state.metrics(),
          sameRoot: state.root === state.edit.textElement,
          focused: document.activeElement === state.root,
          unchanged:
            state.before === new XMLSerializer().serializeToString(state.slide.root.element),
        };
      });
      expect(grown.repeated).toEqual(grown.value);
      expect(grown.sameRoot).toBe(true);
      expect(grown.focused).toBe(true);
      expect(grown.unchanged).toBe(true);
      if (expandsWidth) expect(grown.value.width).toBeGreaterThan(600);
      else expect(grown.value.width).toBe(600);
      if (expandsHeight) expect(grown.value.height).toBeGreaterThan(48);
      else expect(grown.value.height).toBe(48);
      for (const metrics of [grown.beforeUpdate, grown.afterUpdate]) {
        expect(Math.abs(metrics.width - grown.value.width * 0.75)).toBeLessThan(0.75);
        expect(Math.abs(metrics.height - grown.value.height * 0.75)).toBeLessThan(0.75);
        expect(metrics.fontSize).toBe(grown.initial.fontSize);
        expect(metrics.glyphHeight).toBeCloseTo(grown.initial.glyphHeight, 1);
      }
      const roundtrip = await state.evaluate(async (state) => {
        const {
          editSlideElements,
          readSlideElements,
          refreshSlideParts,
          renderEditableText,
          renderSlide,
        } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
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
        state.presentation = refreshSlideParts(state.presentation, state.files, [0]);
        state.slide = state.presentation.slides[0];
        const reloaded = readSlideElements(state.presentation, state.slide)[0];
        const preview = renderSlide(state.presentation, state.slide, { onNodeRendered() {} });
        state.host.append(preview.element);
        await preview.ready;
        const previewNode = preview.element.querySelector('[data-node-id="2"]');
        const rect = previewNode.getBoundingClientRect();
        preview.dispose();
        preview.element.remove();
        state.edit.dispose();
        state.edit.element.remove();
        state.edit = renderEditableText(state.presentation, state.slide, reloaded);
        state.host.append(state.edit.element);
        await state.edit.ready;
        state.root = state.edit.textElement;
        return {
          reloaded,
          reentered: state.edit.extract(),
          previewWidth: rect.width,
          previewHeight: rect.height,
          otherWidth: state.slide.nodes.find((node) => node.id === '3').size.w,
        };
      });
      expect(roundtrip.reloaded.width).toBeCloseTo(grown.value.width, 2);
      expect(roundtrip.reloaded.height).toBeCloseTo(grown.value.height, 2);
      expect(roundtrip.reentered.width).toBeCloseTo(grown.value.width, 2);
      expect(roundtrip.reentered.height).toBeCloseTo(grown.value.height, 2);
      expect(roundtrip.previewWidth).toBeCloseTo(grown.value.width * 0.75, 1);
      expect(roundtrip.previewHeight).toBeCloseTo(grown.value.height * 0.75, 1);
      expect(roundtrip.otherWidth).toBe(600);
      await state.evaluate(async (state) => {
        const { createEditableTextRange } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        state.root.focus();
        const range = createEditableTextRange(state.root, 0, 5, 1, 11);
        document.getSelection().removeAllRanges();
        document.getSelection().addRange(range);
      });
      await page.keyboard.press('Backspace');
      const reduced = await state.evaluate(async (state) => {
        const { editSlideElements, readSlideElements, refreshSlideParts } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        let value = state.edit.extract();
        state.edit.update(value);
        await state.edit.ready;
        value = state.edit.extract();
        const previous = readSlideElements(state.presentation, state.slide);
        const part = editSlideElements(
          state.slide,
          previous.map((item) => (item.nodeId === value.nodeId ? value : item)),
          previous,
          null,
          new Map(),
        );
        state.files.slides.set(state.slide.slidePath, part.xml);
        const refreshed = refreshSlideParts(state.presentation, state.files, [0]);
        return {
          value,
          repeated: state.edit.extract(value),
          reloaded: readSlideElements(refreshed, refreshed.slides[0])[0],
        };
      });
      expect(reduced.repeated).toEqual(reduced.value);
      expect(reduced.reloaded.width).toBeCloseTo(reduced.value.width, 2);
      expect(reduced.reloaded.height).toBeCloseTo(reduced.value.height, 2);
      if (expandsWidth) expect(reduced.value.width).toBeLessThan(grown.value.width);
      else expect(reduced.value.width).toBe(600);
      if (expandsHeight) expect(reduced.value.height).toBeLessThan(grown.value.height);
      else expect(reduced.value.height).toBe(48);
    } finally {
      await state.evaluate(({ edit, host }) => {
        edit.dispose();
        host.remove();
      });
      await state.dispose();
    }
  });
}
