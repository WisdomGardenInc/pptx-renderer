import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const LONG_TEXT = '文字自动调整测试 Auto fit text '.repeat(12);

async function deck() {
  const zip = new JSZip();
  zip.file(
    'ppt/presentation.xml',
    `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`,
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1714500" cy="571500"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

for (const [mode, label, tag] of [
  ['none', '不自动调整', 'noAutofit'],
  ['shrink', '溢出时缩排文字', 'normAutofit'],
  ['resize', '根据文字调整形状大小', 'spAutoFit'],
] as const) {
  for (const wrap of [true, false]) {
    test(`${mode} with wrap=${wrap} survives save and reload`, async ({ page }, testInfo) => {
      const buffer = await deck();
      await page.goto('/');
      await page.locator('input[type="file"]').setInputFiles({
        name: 'text-layout.pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        buffer,
      });
      await expect(page.locator('[data-node-id="2"]')).toBeVisible();
      await page.evaluate(() => window.__pptxHost.select('2'));
      await page.getByRole('radio', { name: label, exact: true }).check();
      await page.getByRole('checkbox', { name: '形状中的文字自动换行' }).setChecked(wrap);
      await page.evaluate((text) => window.__pptxHost.setSelectedText(text), LONG_TEXT);
      await expect(page.getByRole('radio', { name: label, exact: true })).toBeChecked();
      if (mode === 'shrink') {
        await expect
          .poll(() =>
            page.locator('[data-node-id="2"] [data-pptx-text]').getAttribute('data-pptx-fit-scale'),
          )
          .not.toBeNull();
      }
      const result = await page.evaluate(() => {
        const host = window.__pptxHost;
        const shape = host.selectedShape;
        const text = document.querySelector('[data-node-id="2"] [data-pptx-text]');
        const properties = shape.textBody.bodyProperties;
        return {
          size: { ...shape.size },
          wrap: properties.attr('wrap'),
          modes: properties.allChildren().map((child) => child.localName),
          whiteSpace: getComputedStyle(text).whiteSpace,
          scale: Number(text.getAttribute('data-pptx-fit-scale') ?? 1),
          part: host.serializeCurrentSlide(),
        };
      });
      expect(result.wrap).toBe(wrap ? 'square' : 'none');
      expect(result.modes).toEqual([tag]);
      expect(result.whiteSpace).toBe(wrap ? 'normal' : 'nowrap');
      if (mode === 'resize') {
        if (wrap) {
          expect(result.size.w).toBe(180);
          expect(result.size.h).toBeGreaterThan(60);
        } else {
          expect(result.size.w).toBeGreaterThan(180);
        }
      } else {
        expect(result.size).toEqual({ w: 180, h: 60 });
        if (mode === 'none') expect(result.scale).toBe(1);
        else expect(result.scale).toBeLessThan(1);
      }
      await page.screenshot({ path: testInfo.outputPath('text-layout.png') });
      await page.evaluate(
        async ({ bytes, part }) => {
          const host = window.__pptxHost;
          await host.loadFile(new Uint8Array(bytes).buffer, [part]);
          host.select('2');
        },
        { bytes: [...buffer], part: result.part },
      );
      await expect(page.getByRole('radio', { name: label, exact: true })).toBeChecked();
      await expect(page.getByRole('checkbox', { name: '形状中的文字自动换行' })).toBeChecked({
        checked: wrap,
      });
      const reloaded = await page.evaluate(() => ({ ...window.__pptxHost.selectedShape.size }));
      expect(reloaded.w).toBeCloseTo(result.size.w, 2);
      expect(reloaded.h).toBeCloseTo(result.size.h, 2);
      if (mode === 'resize') {
        await page.evaluate(() => window.__pptxHost.setSelectedText('Short'));
        const smaller = await page.evaluate(() => ({ ...window.__pptxHost.selectedShape.size }));
        expect(wrap ? smaller.h : smaller.w).toBeLessThan(wrap ? reloaded.h : reloaded.w);
      }
    });
  }
}
