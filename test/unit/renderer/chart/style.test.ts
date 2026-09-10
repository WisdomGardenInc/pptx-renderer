import { describe, expect, it } from 'vitest';
import {
  extractChartLineStyle,
  extractSeriesColor,
  extractSeriesLineNoFill,
  extractSeriesLineWidth,
  markerSizeToPx,
} from '../../../../src/renderer/chart/style';
import { parseXml } from '../../../../src/parser/XmlParser';
import { createMockRenderContext } from '../../helpers/mockContext';

describe('chart style helpers', () => {
  it('converts OOXML line width and marker size into renderer pixels', () => {
    const series = parseXml(`
      <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:spPr><a:ln w="25400"/></c:spPr>
      </c:ser>
    `);

    expect(extractSeriesLineWidth(series)).toBe(2);
    expect(markerSizeToPx(9)).toBe(12);
  });

  it('detects line noFill and suppresses chart line style', () => {
    const series = parseXml(`
      <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
      </c:ser>
    `);
    const line = series.child('spPr').child('ln');

    expect(extractSeriesLineNoFill(series)).toBe(true);
    expect(extractChartLineStyle(line, createMockRenderContext())).toBeUndefined();
  });

  describe('series picture fill', () => {
    const CHART_PATH = 'ppt/charts/chart1.xml';
    const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

    function pictureFilledSeries() {
      return parseXml(`
        <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:spPr>
            <a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
            <a:ln><a:noFill/></a:ln>
          </c:spPr>
          <c:pictureOptions><c:pictureFormat val="stack"/></c:pictureOptions>
        </c:ser>
      `);
    }

    function ctxWithChartImage(overrides?: { target?: string; targetMode?: string }) {
      const ctx = createMockRenderContext();
      ctx.partPath = CHART_PATH;
      ctx.presentation.media.set('ppt/media/image19.png', PNG);
      ctx.presentation.chartRels = new Map([
        [
          CHART_PATH,
          new Map([
            [
              'rId3',
              {
                type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                target: overrides?.target ?? '../media/image19.png',
                ...(overrides?.targetMode ? { targetMode: overrides.targetMode } : {}),
              },
            ],
          ]),
        ],
      ]);
      return ctx;
    }

    it('turns a blipFill series into a repeating pattern instead of a palette color', () => {
      const ctx = ctxWithChartImage();
      const fill = extractSeriesColor(pictureFilledSeries(), ctx) as {
        image: string;
        repeat: string;
      };

      expect(fill).toBeDefined();
      expect(fill.repeat).toBe('repeat');
      expect(fill.image.startsWith('blob:')).toBe(true);
      // Reuses the shared media cache so the same part is not re-blobbed.
      expect(ctx.mediaUrlCache.get('ppt/media/image19.png')).toBe(fill.image);
    });

    it('prefers an explicit solidFill over the picture fill', () => {
      const ser = parseXml(`
        <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:spPr>
            <a:solidFill><a:srgbClr val="AB0515"/></a:solidFill>
            <a:blipFill><a:blip r:embed="rId3"/></a:blipFill>
          </c:spPr>
        </c:ser>
      `);

      expect(extractSeriesColor(ser, ctxWithChartImage())).toBe('#AB0515');
    });

    it('falls back to the palette when the picture cannot be resolved', () => {
      const ser = pictureFilledSeries();

      const noPartPath = ctxWithChartImage();
      noPartPath.partPath = undefined;
      expect(extractSeriesColor(ser, noPartPath)).toBeUndefined();

      const missingRel = ctxWithChartImage();
      missingRel.presentation.chartRels = new Map();
      expect(extractSeriesColor(ser, missingRel)).toBeUndefined();

      const missingMedia = ctxWithChartImage({ target: '../media/nope.png' });
      expect(extractSeriesColor(ser, missingMedia)).toBeUndefined();

      const external = ctxWithChartImage({ targetMode: 'External' });
      expect(extractSeriesColor(ser, external)).toBeUndefined();
    });
  });
});
