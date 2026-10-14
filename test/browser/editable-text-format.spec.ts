import { expect, test } from '@playwright/test';

const inheritedColorCases = [
  {
    name: 'slide color mapping',
    fill: '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>',
    fontRef: '',
    expected: 'rgb(255, 0, 0)',
  },
  {
    name: 'tinted text',
    fill: '<a:solidFill><a:srgbClr val="000000"><a:tint val="0"/></a:srgbClr></a:solidFill>',
    fontRef: '',
    expected: 'rgb(255, 255, 255)',
  },
  {
    name: 'shape font reference',
    fill: '',
    fontRef: '<a:srgbClr val="FF0000"/>',
    expected: 'rgb(255, 0, 0)',
  },
  {
    name: 'text color transparency',
    fill: '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
    fontRef: '',
    expected: 'rgba(255, 0, 0, 0.5)',
  },
  {
    name: 'explicit run color',
    fill: '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
    fontRef: '<a:srgbClr val="FF0000"/>',
    expected: 'rgb(0, 255, 0)',
  },
];

for (const colorCase of inheritedColorCases) {
  test(`new paragraphs retain ${colorCase.name} through typing and reload`, async ({ page }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(async ({ fill, fontRef }) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const {
        readSlideElements,
        renderSlide,
        renderEditableText,
        createEditableTextRange,
        parseXml,
      } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { presentation, slide, files } = createEditableTextFixture(
        undefined,
        '',
        `<a:r><a:rPr sz="2400">${fill}</a:rPr><a:t>Colored</a:t></a:r>`,
        { empty: true },
      );
      const namespaces =
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
      const mapping = `<p:clrMapOvr ${namespaces}><a:overrideClrMapping tx1="accent1"/></p:clrMapOvr>`;
      slide.root.element.appendChild(parseXml(mapping).element);
      slide.colorMapOverrideMode = 'override';
      slide.colorMapOverride = new Map([['tx1', 'accent1']]);
      presentation.slideToLayout.set(0, 'layout');
      presentation.layoutToMaster.set('layout', 'master');
      presentation.masterToTheme.set('master', 'theme');
      presentation.themes.set('theme', {
        colorScheme: new Map([
          ['dk1', '000000'],
          ['accent1', 'FF0000'],
        ]),
        majorFont: { latin: 'Calibri', ea: '', cs: '' },
        minorFont: { latin: 'Calibri', ea: '', cs: '' },
        fillStyles: [],
        bgFillStyles: [],
        lineStyles: [],
        effectStyles: [],
      });
      if (fontRef)
        slide.nodes[0].source.element.appendChild(
          parseXml(`<p:style ${namespaces}><a:fontRef idx="minor">${fontRef}</a:fontRef></p:style>`)
            .element,
        );
      const element = readSlideElements(presentation, slide)[0];
      const preview = renderSlide(presentation, slide);
      document.body.append(preview.element);
      await preview.ready;
      const originalColor = getComputedStyle(preview.element.querySelector('span')).color;
      const edit = renderEditableText(presentation, slide, element, { editableTextAutoFit: true });
      document.body.append(edit.element);
      await edit.ready;
      const root = edit.textElement;
      const state = {
        presentation,
        slide,
        files,
        edit,
        preview,
        root,
        draft: element,
        originalColor,
      };
      root.addEventListener('beforeinput', (event) => {
        if (event.inputType !== 'insertParagraph') return;
        event.preventDefault();
        state.draft = edit.extract(state.draft);
        const previous = state.draft.paragraphs.at(-1);
        const run = previous.runs.at(-1);
        state.draft = {
          ...state.draft,
          paragraphs: [...state.draft.paragraphs, { ...previous, runs: [{ ...run, text: '' }] }],
        };
        edit.update(state.draft);
        const index = state.draft.paragraphs.length - 1;
        const range = createEditableTextRange(root, index, 0, index, 0);
        document.getSelection().removeAllRanges();
        document.getSelection().addRange(range);
      });
      root.addEventListener('input', () => {
        state.draft = edit.extract(state.draft);
      });
      root.focus();
      const range = createEditableTextRange(root, 0, 7, 0, 7);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      return state;
    }, colorCase);
    try {
      expect(await state.evaluate((state) => state.originalColor)).toBe(colorCase.expected);
      await page.keyboard.press('Enter');
      await page.keyboard.press('Enter');
      await page.keyboard.type('Continued');
      expect(
        await state.evaluate(
          ({ root }) => getComputedStyle(root.children[2].querySelector('span')).color,
        ),
      ).toBe(colorCase.expected);
      const saved = await state.evaluate(async (state) => {
        const { readSlideElements, editSlideElements, refreshSlideParts } =
          await import('/dist/aiden0z-pptx-renderer.browser.es.js');
        const previous = readSlideElements(state.presentation, state.slide);
        const part = editSlideElements(
          state.slide,
          previous.map((item) => (item.nodeId === state.draft.nodeId ? state.draft : item)),
          previous,
          null,
          new Map(),
        );
        state.files.slides.set(state.slide.slidePath, part.xml);
        const presentation = refreshSlideParts(state.presentation, state.files, [0]);
        const element = readSlideElements(presentation, presentation.slides[0])[0];
        state.edit.update(element, { presentation, slide: presentation.slides[0] });
        await state.edit.ready;
        return {
          text: state.root.children[2].textContent,
          color: getComputedStyle(state.root.children[2].querySelector('span')).color,
          focused: document.activeElement === state.root,
        };
      });
      expect(saved).toEqual({ text: 'Continued', color: colorCase.expected, focused: true });
    } finally {
      await state.evaluate(({ edit, preview }) => {
        edit.dispose();
        preview.dispose();
      });
      await state.dispose();
    }
  });
}

for (const textBox of [true, false]) {
  test(`typing into a saved empty paragraph retains its formatting (textBox=${textBox})`, async ({
    page,
  }) => {
    await page.route('https://example.com/background.png', (route) => route.abort());
    await page.goto('/test/browser/blank.html');
    const state = await page.evaluateHandle(async (textBox) => {
      const { createEditableTextFixture } = await import('/test/fixtures/editable-text.ts');
      const {
        editSlideElements,
        readSlideElements,
        refreshSlideParts,
        renderEditableText,
        createEditableTextRange,
      } = await import('/dist/aiden0z-pptx-renderer.browser.es.js');
      const { presentation, slide, files, element } = createEditableTextFixture(undefined, '', '', {
        textBox,
      });
      const styled = {
        ...element.paragraphs[0].runs[1],
        text: '',
        color: '#FF0000',
        fontSize: 32,
        fontName: 'Georgia',
        fontEa: 'Noto Sans CJK TC',
        italic: true,
        underline: true,
      };
      const paragraph = { ...element.paragraphs[0], runs: [styled] };
      const draft = { ...element, paragraphs: [...element.paragraphs, paragraph, paragraph] };
      const previous = readSlideElements(presentation, slide);
      const part = editSlideElements(
        slide,
        previous.map((item) => (item.nodeId === draft.nodeId ? draft : item)),
        previous,
        null,
        new Map(),
      );
      files.slides.set(slide.slidePath, part.xml);
      const refreshed = refreshSlideParts(presentation, files, [0]);
      const reloaded = readSlideElements(refreshed, refreshed.slides[0])[0];
      const edit = renderEditableText(refreshed, refreshed.slides[0], reloaded);
      const host = document.createElement('div');
      Object.assign(host.style, {
        position: 'relative',
        transform: 'scale(0.75)',
        transformOrigin: 'top left',
      });
      host.append(edit.element);
      document.body.append(host);
      await Promise.all([edit.ready, document.fonts.ready]);
      const root = edit.textElement;
      root.focus();
      const range = createEditableTextRange(root, 2, 0, 2, 0);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(range);
      const state = { edit, root, host, draft: reloaded, styled };
      root.addEventListener('input', () => {
        state.draft = edit.extract(state.draft);
      });
      return state;
    }, textBox);
    try {
      await page.keyboard.insertText('Continued');
      const result = await state.evaluate(async (state) => {
        const run = state.draft.paragraphs[2].runs[0];
        state.edit.update(state.draft);
        await state.edit.ready;
        const span = state.root.children[2].querySelector('[data-pptx-run="0"]');
        const computed = getComputedStyle(span);
        const expected = { ...state.styled, text: 'Continued' };
        return {
          run,
          expected,
          color: computed.color,
          fontSize: computed.fontSize,
          fontFamily: computed.fontFamily,
          fontWeight: computed.fontWeight,
          fontStyle: computed.fontStyle,
          decoration: computed.textDecorationLine,
          focused: document.activeElement === state.root,
          text: span.textContent,
        };
      });
      expect(result.run).toEqual(result.expected);
      expect(result.color).toBe('rgb(255, 0, 0)');
      expect(Number.parseFloat(result.fontSize)).toBeCloseTo((32 * 96) / 72, 2);
      expect(result.fontFamily).toContain('Georgia');
      expect(result.fontWeight).toBe('700');
      expect(result.fontStyle).toBe('italic');
      expect(result.decoration).toContain('underline');
      expect(result.text).toBe('Continued');
      expect(result.focused).toBe(true);
    } finally {
      await state.evaluate(({ edit, host }) => {
        edit.dispose();
        host.remove();
      });
      await state.dispose();
    }
  });
}
