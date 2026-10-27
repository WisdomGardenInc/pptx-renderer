import { expect, test } from '@playwright/test';
import { PNG } from 'pngjs';

for (const alpha of [0, 0.5, 1]) {
  test(`alpha-only save and reload preserves fill appearance at ${alpha}`, async ({ page }) => {
    await page.goto('/test/browser/blank.html');
    await page.evaluate(() => {
      document.body.style.margin = '0';
      document.body.style.background = 'white';
    });
    const before = await page.evaluate(async () => {
      const { createFillAlphaFixture } = await import('/test/fixtures/fill-alpha.ts');
      const { renderSlide } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const fixture = createFillAlphaFixture();
      const render = renderSlide(fixture.presentation, fixture.slide, { onNodeRendered() {} });
      document.body.appendChild(render.element);
      await render.ready;
      window.alphaFixture = fixture;
      window.alphaRender = render;
      return fixture.elements.map(({ fillColor, fillKind }) => ({ fillColor, fillKind }));
    });
    const original = PNG.sync.read(await page.screenshot());
    const after = await page.evaluate(async (alpha) => {
      const { editSlideElements, refreshSlideParts, readSlideElements, renderSlide } =
        await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const fixture = window.alphaFixture;
      const next = structuredClone(fixture.elements);
      for (const element of next) element.fillAlpha = alpha;
      const saved = editSlideElements(fixture.slide, next, fixture.elements, null, new Map());
      fixture.files.slides.set(fixture.slide.slidePath, saved.xml);
      const updated = refreshSlideParts(fixture.presentation, fixture.files, [0]);
      window.alphaRender.dispose();
      window.alphaRender.element.remove();
      const rendered = renderSlide(updated, updated.slides[0], { onNodeRendered() {} });
      document.body.appendChild(rendered.element);
      await rendered.ready;
      const gradient = rendered.element.querySelector('[data-node-id="2"] linearGradient');
      return {
        colors: readSlideElements(updated, updated.slides[0]).map(({ fillColor, fillKind }) => ({
          fillColor,
          fillKind,
        })),
        stops: [...gradient.querySelectorAll('stop')].map((stop) =>
          stop.getAttribute('stop-color'),
        ),
        gradientFill: rendered.element
          .querySelector('[data-node-id="2"] path')
          .getAttribute('fill'),
      };
    }, alpha);
    expect(after.colors).toEqual(
      before.map((style) => ({
        ...style,
        fillKind: alpha < 1 && style.fillKind === 'theme' ? 'solid' : style.fillKind,
      })),
    );
    expect(after.gradientFill).toMatch(/^url\(#/);
    expect(after.stops).toHaveLength(2);
    const edited = PNG.sync.read(await page.screenshot());
    // Real browser pixels must preserve RGB and change only compositing over white.
    // Sample both ends of the gradient and all three solid fills, away from strokes.
    for (const x of [15, 135, 225, 425, 625, 825]) {
      const offset = (50 * original.width + x) * 4;
      for (let channel = 0; channel < 3; channel++) {
        const expected = Math.round(original.data[offset + channel] * alpha + 255 * (1 - alpha));
        expect(
          Math.abs(edited.data[offset + channel] - expected),
          `pixel ${x}, channel ${channel}`,
        ).toBeLessThanOrEqual(2);
      }
    }
  });
}
