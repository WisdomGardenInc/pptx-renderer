import { expect, test } from '@playwright/test';

for (const [mode, fit] of [
  ['none', '<a:noAutofit/>'],
  ['shrink', '<a:normAutofit fontScale="80000" lnSpcReduction="20000"/>'],
  ['resize', '<a:spAutoFit/>'],
] as const) {
  for (const anchor of ['t', 'ctr', 'b']) {
    for (const wrap of ['square', 'none']) {
      test(`opening text retains every glyph position (${mode}, ${anchor}, ${wrap})`, async ({
        page,
      }) => {
        await page.route('https://example.com/background.png', (route) => route.abort());
        await page.goto('/test/browser/blank.html');
        const result = await page.evaluate(
          async ({ fit, anchor, wrap }) => {
            const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
            const { renderSlide, renderEditableText } = await import('/src/index.ts');
            const fixture = createEditableTextFixture(
              `<a:bodyPr wrap="${wrap}" anchor="${anchor}" lIns="190500" rIns="95250">${fit}</a:bodyPr>`,
              '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
              '<a:r><a:rPr sz="2400"/><a:t> 文字排版 Alignment and wrapping ' +
                '文字排版 Alignment and wrapping '.repeat(8) +
                '</a:t></a:r>',
              { height: 48 },
            );
            const preview = renderSlide(fixture.presentation, fixture.slide, {
              onNodeRendered() {},
            });
            const edit = renderEditableText(fixture.presentation, fixture.slide, fixture.element, {
              editableTextAutoFit: true,
            });
            const previewHost = document.createElement('div');
            const editHost = document.createElement('div');
            for (const host of [previewHost, editHost]) {
              Object.assign(host.style, {
                position: 'absolute',
                top: '0',
                left: '0',
                whiteSpace: 'pre',
                transform: 'scale(0.75)',
                transformOrigin: 'top left',
              });
              document.body.append(host);
            }
            previewHost.append(preview.element);
            editHost.append(edit.element);
            await Promise.all([preview.ready, edit.ready, document.fonts.ready]);
            await new Promise((resolve) =>
              requestAnimationFrame(() => requestAnimationFrame(resolve)),
            );
            const positions = (root) => {
              const result = [];
              const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
              while (walker.nextNode()) {
                const node = walker.currentNode;
                for (let offset = 0; offset < node.textContent.length; offset++) {
                  const range = document.createRange();
                  range.setStart(node, offset);
                  range.setEnd(node, offset + 1);
                  const rect = range.getBoundingClientRect();
                  result.push(
                    [rect.x, rect.y, rect.width, rect.height].map(
                      (value) => Math.round(value * 10) / 10,
                    ),
                  );
                }
              }
              return result;
            };
            const previewRoot = preview.element.querySelector(
              '[data-node-id="2"] [data-pptx-text]',
            );
            const draft = edit.extract();
            const result = {
              preview: positions(previewRoot),
              editing: positions(edit.textElement),
              previewScale: previewRoot.dataset.pptxFitScale,
              editingScale: edit.textElement.dataset.pptxFitScale,
              width: draft.width,
              height: draft.height,
              originalWidth: fixture.element.width,
            };
            preview.dispose();
            edit.dispose();
            previewHost.remove();
            editHost.remove();
            return result;
          },
          { fit, anchor, wrap },
        );
        expect(result.editingScale).toBe(result.previewScale);
        expect(result.width).toBe(result.originalWidth);
        expect(result.height).toBe(48);
        expect(result.editing.length).toBe(result.preview.length);
        const differences = result.editing.flatMap((position, index) =>
          position.map((value, axis) => Math.abs(value - result.preview[index][axis])),
        );
        expect(Math.max(...differences)).toBeLessThanOrEqual(0.15);
      });
    }
  }
}
