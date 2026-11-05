import { expect, test } from '@playwright/test';

for (const [value, enabled] of [
  ['true', true],
  ['1', true],
  ['false', false],
  ['0', false],
] as const) {
  test(`editor values agree with Chromium for OOXML boolean ${value}`, async ({ page }) => {
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(async (value) => {
      const { createOoxmlEditorFixture } = await import('/test/fixtures/ooxml-editor-values.ts');
      const { renderSlide, editSlideElements, readSlideElements, refreshSlideParts } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const fixture = createOoxmlEditorFixture(value);
      const readDom = async (presentation, slide) => {
        const rendered = renderSlide(presentation, slide, { onNodeRendered() {} });
        document.body.appendChild(rendered.element);
        await rendered.ready;
        const shape = rendered.element.querySelector('[data-node-id="2"]');
        const spans = [...shape.querySelectorAll('span')].filter((span) =>
          ['First', 'Middle', 'Last'].includes(span.textContent),
        );
        const result = spans.map((span) => {
          const style = getComputedStyle(span);
          const paragraph = getComputedStyle(span.parentElement);
          return {
            weight: style.fontWeight,
            italic: style.fontStyle,
            before: paragraph.marginTop,
            after: paragraph.marginBottom,
          };
        });
        rendered.dispose();
        rendered.element.remove();
        return result;
      };
      const original = await readDom(fixture.presentation, fixture.slide);
      const next = structuredClone(fixture.elements);
      next[0].locked = !next[0].locked;
      Object.assign(next[0].paragraphs[1].runs[0], {
        bold: !next[0].paragraphs[1].runs[0].bold,
        italic: !next[0].paragraphs[1].runs[0].italic,
      });
      Object.assign(next[0].paragraphs[1], {
        spaceBefore: 1.5,
        spaceAfter: 0.5,
        spaceAfterIsPt: false,
      });
      const saved = editSlideElements(fixture.slide, next, fixture.elements, null, new Map());
      fixture.files.slides.set(fixture.slide.slidePath, saved.xml);
      const updated = refreshSlideParts(fixture.presentation, fixture.files, [0]);
      const after = readSlideElements(updated, updated.slides[0]);
      const edited = await readDom(updated, updated.slides[0]);
      return { before: fixture.elements[0], original, after: after[0], edited };
    }, value);
    expect(result.before.locked).toBe(enabled);
    expect(result.before.type).toBe(enabled ? 'text' : 'shape');
    expect(result.before.paragraphs[1]).toMatchObject({
      spaceBefore: 2,
      spaceBeforeIsPt: false,
      spaceAfter: 12,
    });
    expect(result.original[1].weight).toBe(enabled ? '700' : '400');
    expect(result.original[1].italic).toBe(enabled ? 'italic' : 'normal');
    expect(parseFloat(result.original[1].before)).toBeCloseTo(76.16, 2);
    expect(parseFloat(result.original[1].after)).toBeCloseTo(16, 2);
    expect(result.original[0].before).toBe('0px');
    expect(result.original[2].after).toBe('0px');
    expect(result.after.locked).toBe(!enabled);
    expect(result.after.paragraphs[1].runs[0]).toMatchObject({ bold: !enabled, italic: !enabled });
    expect(result.edited[1].weight).toBe(enabled ? '400' : '700');
    expect(result.edited[1].italic).toBe(enabled ? 'normal' : 'italic');
    expect(parseFloat(result.edited[1].before)).toBeCloseTo(57.12, 2);
    expect(parseFloat(result.edited[1].after)).toBeCloseTo(19.04, 2);
    expect(result.edited[0]).toEqual(result.original[0]);
    expect(result.edited[2]).toEqual(result.original[2]);
  });
}

for (const defaultIndex of [false, true]) {
  test(`placeholder styles agree between projection, preview and editing overlay (default idx ${defaultIndex})`, async ({
    page,
  }) => {
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(async (defaultIndex) => {
      const { createPlaceholderStyleFixture } =
        await import('/test/fixtures/ooxml-editor-values.ts');
      const { renderSlide, renderEditableText } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const fixture = createPlaceholderStyleFixture(defaultIndex);
      const rendered = renderSlide(fixture.presentation, fixture.slide, { onNodeRendered() {} });
      document.body.appendChild(rendered.element);
      await rendered.ready;
      const overlay = renderEditableText(fixture.presentation, fixture.slide, fixture.elements[0]);
      document.body.appendChild(overlay.element);
      const read = (root) => {
        const span = [...root.querySelectorAll('span')].find(
          (span) => span.textContent === 'Inherited',
        );
        const style = getComputedStyle(span);
        const paragraph = getComputedStyle(span.parentElement);
        return {
          weight: style.fontWeight,
          italic: style.fontStyle,
          size: style.fontSize,
          align: paragraph.textAlign,
        };
      };
      try {
        return {
          element: fixture.elements[0],
          preview: read(rendered.element),
          overlay: read(overlay.element),
        };
      } finally {
        overlay.dispose();
        rendered.dispose();
      }
    }, defaultIndex);
    expect(result.element.paragraphs[0]).toMatchObject({
      align: 'right',
      spaceBefore: 1.5,
      spaceBeforeIsPt: false,
      spaceAfter: 6,
      runs: [{ bold: true, italic: true, fontSize: 28 }],
    });
    expect(result.preview).toEqual(result.overlay);
    expect(result.preview.weight).toBe('700');
    expect(result.preview.italic).toBe('italic');
    expect(parseFloat(result.preview.size)).toBeCloseTo((28 * 4) / 3, 3);
    expect(result.preview.align).toBe('right');
  });
}
