import { expect, test } from '@playwright/test';
import { PNG } from 'pngjs';

const cases = [
  ['dark', ['#000000', '#EEECE1']],
  ['identity', ['#FFFFFF', '#1F497D']],
  ['layout', ['#FFFFFF', '#1F497D']],
  ['inherit-layout', ['#FFFFFF', '#1F497D']],
  ['slide', ['#1F497D', '#FFFFFF']],
] as const;

for (const [mapping, colors] of cases) {
  test(`editor projection matches browser colors before and after alpha edit: ${mapping}`, async ({
    page,
  }) => {
    await page.goto('/test/browser/blank.html');
    const before = await page.evaluate(async (mapping) => {
      const { createColorMapFixture } = await import('/test/fixtures/color-map.ts');
      const { renderSlide } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      document.body.style.margin = '0';
      document.body.style.background = 'white';
      const fixture = createColorMapFixture(mapping);
      const rendered = renderSlide(fixture.presentation, fixture.slide, { onNodeRendered() {} });
      document.body.appendChild(rendered.element);
      await rendered.ready;
      window.colorMapFixture = fixture;
      window.colorMapRender = rendered;
      return {
        fills: fixture.elements.map((element) => element.fillColor),
        bullets: fixture.elements.map((element) => element.paragraphs[0].bullet.color),
        runs: fixture.elements.map((element) => element.paragraphs[0].runs[0].color),
        spans: fixture.elements.map((element) =>
          [
            ...rendered.element
              .querySelector(`[data-node-id="${element.nodeId}"]`)
              .querySelectorAll('span'),
          ].map((span) => getComputedStyle(span).color),
        ),
      };
    }, mapping);
    expect(before.fills).toEqual(colors);
    expect(before.bullets).toEqual([...colors].reverse());
    expect(before.runs).toEqual([...colors].reverse());
    const rgb = colors.map((color) =>
      color
        .slice(1)
        .match(/../g)
        .map((part) => parseInt(part, 16)),
    );
    for (const [index, spans] of before.spans.entries()) {
      expect(spans).toHaveLength(2);
      expect(spans).toEqual(Array(2).fill(`rgb(${rgb[1 - index].join(', ')})`));
    }
    const original = PNG.sync.read(await page.screenshot());
    const after = await page.evaluate(async () => {
      const { editSlideElements, refreshSlideParts, readSlideElements, renderSlide } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { files, presentation, slide, elements } = window.colorMapFixture;
      const next = structuredClone(elements).map((element) => ({ ...element, fillAlpha: 0.5 }));
      const saved = editSlideElements(slide, next, elements, null, new Map());
      files.slides.set(slide.slidePath, saved.xml);
      const updated = refreshSlideParts(presentation, files, [0]);
      window.colorMapRender.dispose();
      window.colorMapRender.element.remove();
      const rendered = renderSlide(updated, updated.slides[0], { onNodeRendered() {} });
      document.body.appendChild(rendered.element);
      await rendered.ready;
      return readSlideElements(updated, updated.slides[0]).map((element) => ({
        fillColor: element.fillColor,
        fillAlpha: element.fillAlpha,
        paragraphs: element.paragraphs,
      }));
    });
    expect(after.map((element) => element.fillColor)).toEqual(colors);
    expect(after.map((element) => element.fillAlpha)).toEqual([0.5, 0.5]);
    expect(after.map((element) => element.paragraphs[0].bullet.color)).toEqual(before.bullets);
    expect(after.map((element) => element.paragraphs[0].runs[0].color)).toEqual(before.runs);
    const edited = PNG.sync.read(await page.screenshot());
    for (const [index, x] of [90, 290].entries()) {
      const offset = (original.width * 170 + x) * 4;
      expect([...original.data.subarray(offset, offset + 3)]).toEqual(rgb[index]);
      for (let channel = 0; channel < 3; channel++) {
        const expected = Math.round(rgb[index][channel] * 0.5 + 127.5);
        expect(Math.abs(edited.data[offset + channel] - expected)).toBeLessThanOrEqual(2);
      }
    }
  });
}
