import { expect, test } from '@playwright/test';

test.use({ channel: process.env.PLAYWRIGHT_CHANNEL });

test('reports shrink and overflow for rendered text bodies', async ({ page }) => {
  await page.goto('/test/browser/blank.html');
  const reports = await page.evaluate(async () => {
    const { textFitFiles } = await import('/test/fixtures/text-fit.ts');
    const { buildPresentation } = await import('/src/model/Presentation.ts');
    const { renderSlide } = await import('/src/renderer/SlideRenderer.ts');
    const { readTextFit } = await import('/src/renderer/textFit.ts');
    const presentation = buildPresentation(textFitFiles());
    const elements = new Map<string, HTMLElement>();
    const handle = renderSlide(presentation, presentation.slides[0], {
      onNodeRendered: (node, element) => elements.set(node.id, element),
    });
    document.body.style.margin = '0';
    document.body.append(handle.element);
    await handle.ready;
    return Object.fromEntries([...elements].map(([id, element]) => [id, readTextFit(element)]));
  });

  expect(reports['2']).toEqual({ scale: 1, overflowsX: false, overflowsY: false });
  expect(reports['3']!.scale).toBeLessThan(1);
  expect(reports['3']!.overflowsY).toBe(false);
  expect(reports['4']!.scale).toBe(1);
  expect(reports['4']!.overflowsY).toBe(true);
});
