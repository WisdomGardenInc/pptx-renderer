import { expect, test } from '@playwright/test';

for (const wrap of ['square', 'none']) {
  test(`fits new and copied text without changing the original slide (${wrap})`, async ({
    page,
  }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(async (wrap) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const { fitSlideText, editSlideElements, readSlideElements, refreshSlideParts, parseXml } =
        await import('/src/index.ts');
      const fixture = createEditableTextFixture();
      const previous = readSlideElements(fixture.presentation, fixture.slide);
      const original = new XMLSerializer().serializeToString(fixture.slide.root.element);
      const source = {
        ...previous[0],
        autoFit: 'resize',
        wrap,
        width: 120,
        height: 30,
        paragraphs: [{ runs: [{ text: 'Autofitting text '.repeat(20), fontSize: 24 }] }],
      };
      const added = { ...source, id: 'added', nodeId: undefined, sourceXml: undefined };
      const copied = { ...source, id: 'copied', nodeId: undefined, copyStyle: 'snapshot' };
      const requested = [...previous, added, copied];
      const input = structuredClone(requested);
      const fitted = await fitSlideText(fixture.presentation, fixture.slide, requested, previous);
      const unchanged =
        original === new XMLSerializer().serializeToString(fixture.slide.root.element);
      const part = editSlideElements(fixture.slide, fitted, previous, null, new Map());
      fixture.files.slides.set(fixture.slide.slidePath, part.xml);
      const reloaded = refreshSlideParts(fixture.presentation, fixture.files, [0]);
      const restored = readSlideElements(reloaded, reloaded.slides[0]);
      return {
        unchanged,
        inputUnchanged: JSON.stringify(input) === JSON.stringify(requested),
        fitted: fitted.slice(-2).map((element) => ({
          id: element.id,
          nodeId: element.nodeId,
          width: element.width,
          height: element.height,
        })),
        restored: restored.map((element) => ({
          nodeId: element.nodeId,
          width: element.width,
          height: element.height,
        })),
        shapeCount: parseXml(part.xml).child('cSld').child('spTree').children('sp').length,
        inputElements: requested.length,
        temporaryInputs: document.querySelectorAll('[data-pptx-editable]').length,
      };
    }, wrap);
    expect(result.unchanged).toBe(true);
    expect(result.inputUnchanged).toBe(true);
    expect(result.shapeCount).toBe(result.inputElements);
    expect(result.temporaryInputs).toBe(0);
    for (const [index, fitted] of result.fitted.entries()) {
      expect(fitted.nodeId).toBeUndefined();
      if (wrap === 'square') {
        expect(fitted.width).toBe(120);
        expect(fitted.height).toBeGreaterThan(30);
      } else {
        expect(fitted.width).toBeGreaterThan(120);
      }
      const restored = result.restored[index + result.restored.length - 2];
      expect(restored.width).toBeCloseTo(fitted.width, 2);
      expect(restored.height).toBeCloseTo(fitted.height, 2);
    }
  });
}

test('fits grouped text while retaining the parent and sibling geometry', async ({ page }) => {
  await page.route('https://example.com/background.png', (route) => route.abort());
  await page.goto('/test/browser/blank.html');
  const result = await page.evaluate(async () => {
    const { createGroupedTextFixture } = await import('/test/fixtures/grouped-text.ts');
    const { fitSlideText, editSlideElements, readSlideElements, refreshSlideParts, sourceElement } =
      await import('/src/index.ts');
    const fixture = createGroupedTextFixture([{}, { width: 900, height: 200 }]);
    const requested = structuredClone(fixture.elements);
    const target = requested.find((element) => element.nodeId === '2');
    Object.assign(target, { autoFit: 'resize', wrap: 'square', height: 30 });
    target.paragraphs[0].runs[0].text = 'Text inside a group '.repeat(16);
    const serializer = new XMLSerializer();
    const original = serializer.serializeToString(fixture.slide.root.element);
    const parent = serializer.serializeToString(
      sourceElement(fixture.slide, '11').getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xfrm',
      )[0],
    );
    const sibling = serializer.serializeToString(sourceElement(fixture.slide, '3'));
    const fitted = await fitSlideText(
      fixture.presentation,
      fixture.slide,
      requested,
      fixture.elements,
    );
    const unchanged = original === serializer.serializeToString(fixture.slide.root.element);
    const part = editSlideElements(fixture.slide, fitted, fixture.elements, null, new Map());
    fixture.files.slides.set(fixture.slide.slidePath, part.xml);
    const presentation = refreshSlideParts(fixture.presentation, fixture.files, [0]);
    const slide = presentation.slides[0];
    const restored = readSlideElements(presentation, slide).find(
      (element) => element.nodeId === '2',
    );
    return {
      unchanged,
      expected: fitted.find((element) => element.nodeId === '2'),
      restored,
      parentUnchanged:
        parent ===
        serializer.serializeToString(
          sourceElement(slide, '11').getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/drawingml/2006/main',
            'xfrm',
          )[0],
        ),
      siblingUnchanged: sibling === serializer.serializeToString(sourceElement(slide, '3')),
      parentName: sourceElement(slide, '2').parentElement.localName,
    };
  });
  expect(result.unchanged).toBe(true);
  expect(result.parentUnchanged).toBe(true);
  expect(result.siblingUnchanged).toBe(true);
  expect(result.parentName).toBe('grpSp');
  expect(result.expected.height).toBeGreaterThan(30);
  expect(result.restored.width).toBeCloseTo(result.expected.width, 2);
  expect(result.restored.height).toBeCloseTo(result.expected.height, 2);
});

test('fixed text modes retain authored dimensions during fitting', async ({ page }) => {
  await page.route('https://example.com/background.png', (route) => route.abort());
  await page.goto('/test/browser/blank.html');
  const result = await page.evaluate(async () => {
    const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
    const { fitSlideText } = await import('/src/index.ts');
    const fixture = createEditableTextFixture();
    const serializer = new XMLSerializer();
    const original = serializer.serializeToString(fixture.slide.root.element);
    const elements = ['none', 'shrink'].map((mode, index) => ({
      ...fixture.element,
      id: `fixed-${index}`,
      nodeId: undefined,
      sourceXml: undefined,
      autoFit: mode,
      width: 120,
      height: 30,
      paragraphs: [{ runs: [{ text: 'Overflowing text '.repeat(20) }] }],
    }));
    const fitted = await fitSlideText(fixture.presentation, fixture.slide, elements, []);
    return {
      unchanged: original === serializer.serializeToString(fixture.slide.root.element),
      sizes: fitted.map((element) => [element.width, element.height]),
    };
  });
  expect(result.unchanged).toBe(true);
  expect(result.sizes).toEqual([
    [120, 30],
    [120, 30],
  ]);
});
