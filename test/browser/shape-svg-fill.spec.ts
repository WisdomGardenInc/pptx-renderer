import { expect, test } from '@playwright/test';
import { PNG } from 'pngjs';

test.use({ channel: process.env.PLAYWRIGHT_CHANNEL });

for (const lazy of [false, true]) {
  test(`SVG shape fills render inside scaled groups with ${lazy ? 'lazy' : 'eager'} media`, async ({
    page,
  }) => {
    await page.goto('/test/browser/blank.html');
    const result = await page.evaluate(async (lazy) => {
      const { presentation } = await import('/test/fixtures/media-group-coverage.ts');
      const { parseXml } = await import('/src/parser/XmlParser.ts');
      const { parseShapeNode } = await import('/src/model/nodes/ShapeNode.ts');
      const { parseGroupNode } = await import('/src/model/nodes/GroupNode.ts');
      const { renderSlide } = await import('/src/renderer/SlideRenderer.ts');
      const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"`;
      const shape = (id: number, x: number, svg: boolean, raster: boolean, size = 20) => `
        <p:sp ${ns}>
          <p:nvSpPr><p:cNvPr id="${id}" name="icon-${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="${x * 9525}" y="0"/><a:ext cx="${size * 9525}" cy="${size * 9525}"/></a:xfrm>
            <a:custGeom><a:pathLst><a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
              <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
              <a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/>
            </a:path></a:pathLst></a:custGeom>
            <a:blipFill><a:blip ${raster ? 'r:embed="rRaster"' : ''}>
              ${svg ? '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rSvg"/></a:ext></a:extLst>' : ''}
            </a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>
            <a:ln><a:noFill/></a:ln>
          </p:spPr>
        </p:sp>`;
      const group = parseGroupNode(
        parseXml(`<p:grpSp ${ns}>
        <p:nvGrpSpPr><p:cNvPr id="1" name="scaled icons"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm>
          <a:off x="190500" y="190500"/><a:ext cx="2857500" cy="952500"/>
          <a:chOff x="0" y="0"/><a:chExt cx="1428750" cy="476250"/>
        </a:xfrm></p:grpSpPr>
        ${shape(2, 0, true, false)}${shape(3, 25, true, true)}${shape(4, 50, false, true)}
      </p:grpSp>`),
      );
      const p = presentation();
      p.slides[0].nodes = [group, parseShapeNode(parseXml(shape(5, 200, true, false, 40)))];
      p.slides[0].rels.set('rSvg', { type: 'image', target: '../media/icon.svg' });
      p.slides[0].rels.set('rRaster', { type: 'image', target: '../media/fallback.png' });
      const canvas = document.createElement('canvas');
      canvas.width = canvas.height = 20;
      canvas.getContext('2d')!.fillStyle = '#ff0000';
      canvas.getContext('2d')!.fillRect(0, 0, 20, 20);
      const media = new Map([
        [
          'ppt/media/icon.svg',
          new TextEncoder().encode(
            '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="#00ff00"/></svg>',
          ),
        ],
        [
          'ppt/media/fallback.png',
          new Uint8Array(await (await fetch(canvas.toDataURL())).arrayBuffer()),
        ],
      ]);
      if (lazy) {
        p.mediaResolver = {
          resolve: async (target) => {
            const mediaPath = target.replace('../', 'ppt/');
            return { mediaPath, data: media.get(mediaPath)! };
          },
        };
      } else p.media = media;
      const handle = renderSlide(p, p.slides[0]);
      document.body.style.margin = '0';
      document.body.replaceChildren(handle.element);
      await handle.ready;
      const images = [...handle.element.querySelectorAll('svg image')];
      for (const image of images) {
        const img = new Image();
        img.src = image.getAttributeNS('http://www.w3.org/1999/xlink', 'href')!;
        await img.decode();
      }
      return images.length;
    }, lazy);
    expect(result).toBe(4);
    const png = PNG.sync.read(await page.locator('body > div').screenshot());
    const pixel = (x: number, y: number) => {
      const offset = (y * png.width + x) * 4;
      return [...png.data.subarray(offset, offset + 4)];
    };
    // Verify actual composited pixels: SVG-only, SVG preferred over PNG,
    // ordinary PNG, and a standalone SVG shape outside the scaled group.
    expect(pixel(30, 30)).toEqual([0, 255, 0, 255]);
    expect(pixel(80, 30)).toEqual([0, 255, 0, 255]);
    expect(pixel(130, 30)).toEqual([255, 0, 0, 255]);
    expect(pixel(210, 10)).toEqual([0, 255, 0, 255]);
  });
}
