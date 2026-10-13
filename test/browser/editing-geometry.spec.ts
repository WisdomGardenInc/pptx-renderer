import { expect, test } from '@playwright/test';

for (const [dx, dy] of [
  [300, 0],
  [0, 300],
  [300, 100],
  [-300, -100],
]) {
  test(`custom endpoint authoring retains the fixed diamond and layout (${dx}, ${dy})`, async ({
    page,
  }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(
      async ([dx, dy]) => {
        const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
        const {
          createShapeElement,
          createLineEndpointGeometry,
          createFilledLineGeometry,
          writeShapeGeometry,
          writeShapePaint,
          refreshSlideParts,
          renderSlide,
        } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const { files, presentation, slide } = createEditableTextFixture();
        const document = slide.root.element.ownerDocument;
        const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
        const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
        let source = createShapeElement(document, { id: '9', connector: true, preset: 'line' });
        slide.root.element.getElementsByTagNameNS(P, 'spTree')[0].appendChild(source);
        const transform = source.getElementsByTagNameNS(A, 'xfrm')[0];
        transform.setAttribute('flipH', dx < 0 ? '1' : '0');
        transform.setAttribute('flipV', dy < 0 ? '1' : '0');
        source.getElementsByTagNameNS(A, 'off')[0].setAttribute('x', '3810000');
        source.getElementsByTagNameNS(A, 'off')[0].setAttribute('y', '1905000');
        const dimensions = { length: 24, width: 24, outlineWidth: 4 };
        const end = createLineEndpointGeometry('diamond', dimensions);
        const snapshots = [];
        for (const opposite of [
          'none',
          'circle',
          'square',
          'arrow',
          'triangle',
          'reversedTriangle',
          'stealth',
          'diamond',
          'none',
        ]) {
          const geometry = createFilledLineGeometry({
            width: Math.abs(dx),
            height: Math.abs(dy),
            strokeWidth: 8,
            start: createLineEndpointGeometry(opposite, dimensions),
            end,
            startInset: 0,
            endInset: 0,
          });
          source = writeShapeGeometry(
            source,
            { kind: 'custom', ...geometry },
            { convertConnector: true, setExtent: true },
          );
          writeShapePaint(source, {
            fill: { color: '#123456' },
            stroke: { color: null, width: 0, headEnd: null, tailEnd: null },
          });
          files.slides.set(
            slide.slidePath,
            new XMLSerializer().serializeToString(slide.root.element),
          );
          const updated = refreshSlideParts(presentation, files, [0]);
          const rendered = renderSlide(updated, updated.slides[0], { onNodeRendered() {} });
          window.document.body.appendChild(rendered.element);
          try {
            await rendered.ready;
            const svg = rendered.element.querySelector('[data-node-id="9"] svg');
            const path = svg.querySelector('path');
            const d = path.getAttribute('d');
            const bounds = svg.getBoundingClientRect();
            snapshots.push({
              contour: d.substring(d.lastIndexOf('M')),
              bounds: [bounds.x, bounds.y, bounds.width, bounds.height],
              markerCount: svg.querySelectorAll('marker').length,
              overflow: getComputedStyle(svg).overflow,
            });
          } finally {
            rendered.dispose();
            rendered.element.remove();
          }
        }
        return snapshots;
      },
      [dx, dy],
    );
    expect(result).toHaveLength(9);
    expect(result[0].contour).toMatch(/^M/);
    expect(result[0].overflow).toBe('visible');
    for (const snapshot of result) {
      expect(snapshot).toEqual(result[0]);
      expect(snapshot.markerCount).toBe(0);
    }
  });
}
