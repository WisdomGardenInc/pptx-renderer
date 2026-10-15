import { expect, test } from '@playwright/test';

for (const [mode, tag, enabled] of [
  ['legacy', 'noAutofit', false],
  ['none', 'noAutofit', true],
  ['shrink', 'normAutofit', true],
  ['resize', 'spAutoFit', true],
] as const) {
  for (const wrap of ['square', 'none']) {
    for (const offset of [0, 8, 12]) {
      test(`native spaces and undo survive autofit (${mode}, wrap=${wrap}, offset=${offset})`, async ({
        page,
      }) => {
        await page.route('https://example.com/background.png', (route) => route.abort());
        await page.goto('/test/browser/blank.html');
        const state = await page.evaluateHandle(
          async ({ tag, wrap, offset, enabled }) => {
            const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
            const { renderEditableText, createEditableTextRange } = await import('/src/index.ts');
            const fixture = createEditableTextFixture(
              `<a:bodyPr wrap="${wrap}"><a:${tag}/></a:bodyPr>`,
            );
            const edit = renderEditableText(fixture.presentation, fixture.slide, fixture.element, {
              editableTextAutoFit: enabled,
            });
            const host = document.createElement('div');
            Object.assign(host.style, {
              position: 'relative',
              whiteSpace: 'pre',
              transform: 'scale(0.75)',
              transformOrigin: 'top left',
            });
            host.append(edit.element);
            document.body.append(host);
            await edit.ready;
            const root = edit.textElement;
            root.focus();
            const range = createEditableTextRange(root, 0, offset, 0, offset);
            document.getSelection().removeAllRanges();
            document.getSelection().addRange(range);
            return { ...fixture, edit, root, host };
          },
          { tag, wrap, offset, enabled },
        );
        const text = () =>
          state.evaluate(({ edit }) =>
            edit
              .extract()
              .paragraphs[0].runs.map((run) => run.text)
              .join(''),
          );
        const original = 'Hello  world';
        const expected = original.slice(0, offset) + ' a  b' + original.slice(offset);
        try {
          // Separate key events exercise the browser's temporary trailing NBSPs.
          await page.keyboard.type(' a  b');
          expect(await text()).toBe(expected);
          await page.keyboard.press('ControlOrMeta+z');
          expect(await text()).toBe('Hello  world');
          await page.keyboard.press('ControlOrMeta+Shift+z');
          expect(await text()).toBe(expected);
          const saved = await state.evaluate(async (state) => {
            const { editSlideElements, readSlideElements, refreshSlideParts } =
              await import('/src/index.ts');
            const draft = state.edit.extract();
            const previous = readSlideElements(state.presentation, state.slide);
            const part = editSlideElements(
              state.slide,
              previous.map((element) => (element.nodeId === draft.nodeId ? draft : element)),
              previous,
              null,
              new Map(),
            );
            state.files.slides.set(state.slide.slidePath, part.xml);
            const presentation = refreshSlideParts(state.presentation, state.files, [0]);
            const slide = presentation.slides[0];
            const element = readSlideElements(presentation, slide)[0];
            state.edit.update(element, { presentation, slide });
            await state.edit.ready;
            return {
              text: state.edit
                .extract()
                .paragraphs[0].runs.map((run) => run.text)
                .join(''),
              sameRoot: state.edit.textElement === state.root,
              focused: document.activeElement === state.root,
              bold: element.paragraphs[0].runs.at(-1).bold,
            };
          });
          expect(saved).toEqual({
            text: expected,
            sameRoot: true,
            focused: true,
            bold: true,
          });
        } finally {
          await state.evaluate(({ edit, host }) => {
            edit.dispose();
            host.remove();
          });
          await state.dispose();
        }
      });
    }
  }
}
