import { expect, test } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

test('benchmarks projection and editing with increasing numbers of sibling shapes', async ({
  page,
}) => {
  await page.goto('/test/browser/blank.html');
  const result = await page.evaluate(async () => {
    const { createManyShapesFixture } = await import('/test/fixtures/many-shapes.ts');
    const { readSlideElements, renderEditableText } = await import('/src/index.ts');
    const median = (samples) => [...samples].sort((a, b) => a - b)[Math.floor(samples.length / 2)];
    const results = [];
    for (const count of [51, 101, 201, 401]) {
      const fixture = createManyShapesFixture(count);
      const before = new XMLSerializer().serializeToString(fixture.slide.root.element);
      const handle = renderEditableText(fixture.presentation, fixture.slide, fixture.element);
      document.body.appendChild(handle.element);
      await handle.ready;
      const input = handle.textElement;
      const draft = structuredClone(fixture.element);
      const reads = [];
      const updates = [];
      try {
        for (let iteration = 0; iteration < 10; iteration++) {
          readSlideElements(fixture.presentation, fixture.slide);
          handle.update(draft);
        }
        for (let iteration = 0; iteration < 7; iteration++) {
          const start = performance.now();
          const elements = readSlideElements(fixture.presentation, fixture.slide);
          reads.push(performance.now() - start);
          if (elements.length !== count) throw new Error('Incomplete projection');
        }
        for (let iteration = 0; iteration < 15; iteration++) {
          draft.paragraphs[0].runs[0].text = `Typed ${iteration}`;
          const start = performance.now();
          handle.update(draft);
          updates.push(performance.now() - start);
          await handle.ready;
        }
        results.push({
          count,
          readMedianMs: median(reads),
          updateMedianMs: median(updates),
          reads,
          updates,
          stableInput: input === handle.textElement,
          text: handle.extract().paragraphs[0].runs[0].text,
          unchanged: before === new XMLSerializer().serializeToString(fixture.slide.root.element),
        });
      } finally {
        handle.dispose();
        handle.element.remove();
      }
    }
    return { browser: navigator.userAgent, results };
  });
  console.table(
    result.results.map(({ count, readMedianMs, updateMedianMs }) => ({
      count,
      readMedianMs,
      updateMedianMs,
    })),
  );
  if (process.env.PPTX_PERF_REPORT) {
    await mkdir(path.dirname(process.env.PPTX_PERF_REPORT), { recursive: true });
    await writeFile(process.env.PPTX_PERF_REPORT, JSON.stringify(result, null, 2));
  }
  for (const row of result.results) {
    expect(row.stableInput).toBe(true);
    expect(row.unchanged).toBe(true);
    expect(row.text).toBe('Typed 14');
  }
});
