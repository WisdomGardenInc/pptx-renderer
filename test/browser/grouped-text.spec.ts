import { expect, test } from '@playwright/test';

for (const groups of [
  [{}],
  [{ width: 900, height: 200 }],
  [{ rotation: 45, flipH: true, flipV: true }],
  [
    { rotation: 90, width: 900, height: 200 },
    { rotation: 90, flipH: true },
  ],
  [
    { rotation: 30, flipH: true },
    { rotation: -15, flipV: true, width: 300 },
  ],
]) {
  test(`grouped text overlay matches preview and persists typing: ${JSON.stringify(groups)}`, async ({
    page,
  }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(async (groups) => {
      const { createGroupedTextFixture } = await import('/test/fixtures/grouped-text.ts');
      const { renderSlide, renderEditableText, createEditableTextRange } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const fixture = createGroupedTextFixture(groups);
      const before = new XMLSerializer().serializeToString(fixture.slide.root.element);
      const preview = renderSlide(fixture.presentation, fixture.slide, { onNodeRendered() {} });
      const edit = renderEditableText(fixture.presentation, fixture.slide, fixture.element);
      const host = document.createElement('div');
      Object.assign(host.style, {
        position: 'relative',
        transform: 'scale(0.75)',
        transformOrigin: 'top left',
      });
      host.append(preview.element, edit.element);
      document.body.append(host);
      await Promise.all([preview.ready, edit.ready, document.fonts.ready]);
      const rect = (element) => {
        const bounds = element.getBoundingClientRect();
        return { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height };
      };
      const native = preview.element.querySelector('[data-node-id="2"]');
      const overlay = edit.element.querySelector('[data-node-id="2"]');
      const glyph = (element) => {
        const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
        let text = walker.nextNode();
        while (text && !text.textContent.trim()) text = walker.nextNode();
        const range = document.createRange();
        range.setStart(text, 0);
        range.setEnd(text, 1);
        return rect(range);
      };
      const alignment = {
        preview: rect(native),
        overlay: rect(overlay),
        previewGlyph: glyph(native),
        overlayGlyph: glyph(overlay),
      };
      native.style.visibility = 'hidden';
      const root = edit.textElement;
      root.focus();
      const range = createEditableTextRange(root, 0, 12, 0, 12);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      return { fixture, before, preview, edit, host, alignment };
    }, groups);
    try {
      const alignment = await state.evaluate((state) => state.alignment);
      for (const key of ['x', 'y', 'width', 'height']) {
        expect(alignment.overlay[key]).toBeCloseTo(alignment.preview[key], 1);
        expect(alignment.overlayGlyph[key]).toBeCloseTo(alignment.previewGlyph[key], 1);
      }
      await page.keyboard.insertText(' edited');
      const result = await state.evaluate(async (state) => {
        const { readSlideElements, editSlideElements, refreshSlideParts, renderEditableText } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const draft = state.edit.extract();
        state.edit.update(draft);
        const stableRoot = document.activeElement === state.edit.textElement;
        const { fixture } = state;
        const result = editSlideElements(
          fixture.slide,
          fixture.elements.map((element) => (element.nodeId === draft.nodeId ? draft : element)),
          fixture.elements,
          null,
          new Map(),
        );
        fixture.files.slides.set(fixture.slide.slidePath, result.xml);
        const presentation = refreshSlideParts(fixture.presentation, fixture.files, [0]);
        const slide = presentation.slides[0];
        const reloaded = readSlideElements(presentation, slide).find(
          (element) => element.nodeId === '2',
        );
        const edit = renderEditableText(presentation, slide, reloaded);
        state.host.append(edit.element);
        await edit.ready;
        const text = edit.textElement.textContent;
        edit.dispose();
        edit.element.remove();
        return { text, stableRoot, parent: reloaded.parentNodeId, xml: result.xml };
      });
      expect(result.stableRoot).toBe(true);
      expect(result.text).toContain('world edited');
      expect(result.parent).toBe(String(9 + groups.length));
      expect(result.xml).toContain('rad="123"');
    } finally {
      await state.evaluate((state) => {
        state.edit.dispose();
        state.preview.dispose();
        state.host.remove();
      });
      await state.dispose();
    }
  });
}
