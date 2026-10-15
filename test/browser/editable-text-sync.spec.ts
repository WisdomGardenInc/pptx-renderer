import { expect, test } from '@playwright/test';

for (const [mode, tag] of [
  ['none', 'noAutofit'],
  ['shrink', 'normAutofit'],
  ['resize', 'spAutoFit'],
] as const) {
  for (const wrap of ['square', 'none']) {
    for (const offset of [0, 8, 12]) {
      for (const spaces of [1, 2]) {
        test(`spaces survive model synchronization before the next key (${mode}, ${wrap}, offset=${offset}, spaces=${spaces})`, async ({
          page,
        }) => {
          await page.route('https://example.com/background.png', (route) => route.abort());
          await page.goto('/test/browser/blank.html');
          const state = await page.evaluateHandle(
            async ({ tag, wrap, offset }) => {
              const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
              const { renderEditableText, createEditableTextRange } = await import('/src/index.ts');
              const fixture = createEditableTextFixture(
                `<a:bodyPr wrap="${wrap}"><a:${tag}/></a:bodyPr>`,
              );
              const edit = renderEditableText(
                fixture.presentation,
                fixture.slide,
                fixture.element,
                {
                  editableTextAutoFit: true,
                },
              );
              document.body.append(edit.element);
              await edit.ready;
              const root = edit.textElement;
              root.focus();
              const caret = createEditableTextRange(root, 0, offset, 0, offset);
              document.getSelection().removeAllRanges();
              document.getSelection().addRange(caret);
              return { edit, root };
            },
            { tag, wrap, offset },
          );
          try {
            const advance = async () =>
              state.evaluate(async ({ edit, root }) => {
                const caret = document.getSelection().getRangeAt(0).cloneRange();
                caret.setStart(caret.startContainer, caret.startOffset - 1);
                const before = caret.getBoundingClientRect().width;
                edit.update(edit.extract());
                await edit.ready;
                const restored = document.getSelection().getRangeAt(0).cloneRange();
                restored.setStart(restored.startContainer, restored.startOffset - 1);
                return {
                  before,
                  after: restored.getBoundingClientRect().width,
                  sameRoot: edit.textElement === root,
                  focused: document.activeElement === root,
                };
              });
            for (let index = 0; index < spaces; index++) {
              await page.keyboard.press('Space');
              const advanceWidth = await advance();
              expect(advanceWidth.after).toBeGreaterThan(0);
              expect(advanceWidth.after).toBeCloseTo(advanceWidth.before, 1);
              expect(advanceWidth.sameRoot).toBe(true);
              expect(advanceWidth.focused).toBe(true);
            }
            await page.keyboard.type('word');
            const original = 'Hello  world';
            expect(
              await state.evaluate(({ edit }) =>
                edit
                  .extract()
                  .paragraphs[0].runs.map((run) => run.text)
                  .join(''),
              ),
            ).toBe(
              original.slice(0, offset) + ' '.repeat(spaces) + 'word' + original.slice(offset),
            );
          } finally {
            await state.evaluate(({ edit }) => {
              edit.dispose();
              edit.element.remove();
            });
            await state.dispose();
          }
        });
      }
    }
  }
}
