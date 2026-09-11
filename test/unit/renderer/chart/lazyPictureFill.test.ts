import { describe, expect, it, vi } from 'vitest';
import { parseXml } from '../../../../src/parser/XmlParser';
import { extractSeriesColor } from '../../../../src/renderer/chart/style';
import { prefetchChartPictureMedia } from '../../../../src/utils/media';
import type { MediaResolver, ResolvedMedia } from '../../../../src/utils/media';
import type { RelEntry } from '../../../../src/parser/RelParser';
import { createMockRenderContext } from '../../helpers/mockContext';

/**
 * A picture-filled chart series resolves its image while the ECharts option is
 * assembled, and that path is synchronous. Under `lazyMedia` the bytes only
 * exist behind the resolver, so without a prefetch the series silently falls
 * back to a palette color — the demo renders charts with `lazyMedia: true`,
 * which is how this escaped the eager-media unit tests and the oracle suite.
 */
describe('chart picture fills under lazy media', () => {
  const CHART_PATH = 'ppt/charts/chart3.xml';
  const IMAGE_PATH = 'ppt/media/image19.png';
  const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

  const IMAGE_REL_TYPE =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

  function pictureFilledSeries() {
    return parseXml(`
      <c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:spPr>
          <a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
        </c:spPr>
        <c:pictureOptions><c:pictureFormat val="stack"/></c:pictureOptions>
      </c:ser>
    `);
  }

  /** Chart rels as the real deck has them: one image plus parts that are not media. */
  function chartRels(): Map<string, Map<string, RelEntry>> {
    return new Map([
      [
        CHART_PATH,
        new Map<string, RelEntry>([
          ['rId1', { type: 'http://…/relationships/chartStyle', target: 'style1.xml' }],
          ['rId2', { type: 'http://…/relationships/chartColorStyle', target: 'colors1.xml' }],
          ['rId3', { type: IMAGE_REL_TYPE, target: '../media/image19.png' }],
          [
            'rId5',
            {
              type: 'http://…/relationships/package',
              target: '../embeddings/Microsoft_Excel_Worksheet2.xlsx',
            },
          ],
        ]),
      ],
    ]);
  }

  /** Lazy resolver: `media` starts empty and is filled only on resolve, as ZipParser's is. */
  function lazyResolver(media: Map<string, Uint8Array>): MediaResolver {
    return {
      resolve: vi.fn(async (target: string): Promise<ResolvedMedia | undefined> => {
        if (!target.endsWith('image19.png')) return undefined;
        media.set(IMAGE_PATH, PNG);
        return { mediaPath: IMAGE_PATH, data: PNG };
      }),
    };
  }

  function lazyCtx() {
    const ctx = createMockRenderContext();
    ctx.partPath = CHART_PATH;
    ctx.presentation.chartRels = chartRels();
    // Nothing preloaded: lazyMedia keeps every part behind the resolver.
    ctx.presentation.media.clear();
    ctx.presentation.mediaResolver = lazyResolver(ctx.presentation.media);
    return ctx;
  }

  it('falls back to no fill when the picture is still behind the resolver', () => {
    const ctx = lazyCtx();

    // Documents the failure mode the prefetch exists to prevent.
    expect(extractSeriesColor(pictureFilledSeries(), ctx)).toBeUndefined();
  });

  it('paints the picture once the chart media is prefetched', async () => {
    const ctx = lazyCtx();

    await prefetchChartPictureMedia(ctx.presentation.chartRels, ctx.presentation.mediaResolver);

    const fill = extractSeriesColor(pictureFilledSeries(), ctx) as {
      image: string;
      repeat: string;
      pictureFormat: string;
    };

    expect(fill).toBeDefined();
    expect(fill.image.startsWith('blob:')).toBe(true);
    expect(fill.pictureFormat).toBe('stack');
  });

  it('only fetches image relationships, not chart styles or embedded workbooks', async () => {
    const ctx = lazyCtx();
    const resolver = ctx.presentation.mediaResolver as MediaResolver & {
      resolve: ReturnType<typeof vi.fn>;
    };

    await prefetchChartPictureMedia(ctx.presentation.chartRels, resolver);

    expect(resolver.resolve).toHaveBeenCalledTimes(1);
    expect(resolver.resolve).toHaveBeenCalledWith('../media/image19.png');
  });

  it('skips external images and tolerates a resolver that rejects', async () => {
    const rels = new Map([
      [
        CHART_PATH,
        new Map<string, RelEntry>([
          [
            'rId3',
            { type: IMAGE_REL_TYPE, target: 'https://example.com/x.png', targetMode: 'External' },
          ],
          ['rId4', { type: IMAGE_REL_TYPE, target: '../media/image20.png' }],
        ]),
      ],
    ]);
    const resolve = vi.fn(async () => {
      throw new Error('zip entry unreadable');
    });

    // A failed picture must not reject the load: the chart still renders.
    await expect(prefetchChartPictureMedia(rels, { resolve })).resolves.toBeUndefined();
    expect(resolve).toHaveBeenCalledTimes(1);
    expect(resolve).toHaveBeenCalledWith('../media/image20.png');
  });

  it('does nothing without a resolver, keeping eager media untouched', async () => {
    await expect(prefetchChartPictureMedia(chartRels(), undefined)).resolves.toBeUndefined();
  });
});
