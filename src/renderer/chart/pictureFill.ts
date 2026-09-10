/**
 * Geometry for OOXML picture fills (`a:blipFill`) on bar and column series.
 *
 * PowerPoint fits the picture to the bar. `c:pictureFormat` says how: `stretch`
 * (the default) scales one copy over it, while `stack`/`stackScale` tile the
 * picture along it, scaled so the edge across the bar matches its thickness —
 * decks pair horizontal bars with a wide strip and columns with a tall one, so
 * which edge that is follows the bar direction.
 *
 * An ECharts pattern is shape-independent: it tiles at the image's natural pixel
 * size from the canvas origin, cropping the artwork at an arbitrary phase on
 * every bar. Bar rectangles only exist after layout, so this runs once the chart
 * has rendered and rewrites each data point with its own pattern transform.
 *
 * Bubble series take a different route: their fill becomes an ECharts image
 * symbol, which already scales to the bubble.
 */

import type * as EChartsTypes from 'echarts';
import type { EChartsType } from 'echarts/core';

/** Pattern fill produced by `extractSeriesColor` for an `a:blipFill` series. */
export interface PicturePattern {
  image: string;
  repeat: 'repeat' | 'no-repeat';
  /** OOXML `c:pictureFormat`: whether the picture tiles along the bar or fills it once. */
  pictureFormat?: 'stretch' | 'stack' | 'stackScale';
  scaleX?: number;
  scaleY?: number;
  x?: number;
  y?: number;
}

type SeriesRecord = Record<string, unknown>;

/** Charts past this many marks keep the plain tiling; per-point patterns would not pay off. */
const MAX_PATTERNED_POINTS = 2000;

export function isPicturePattern(value: unknown): value is PicturePattern {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as PicturePattern;
  return (
    typeof candidate.image === 'string' &&
    (candidate.repeat === 'repeat' || candidate.repeat === 'no-repeat')
  );
}

/** Read an ECharts `'13%'`-style ratio, defaulting when absent or malformed. */
export function parsePercentRatio(value: unknown, fallback: number): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string') return fallback;
  const match = /^(-?\d+(?:\.\d+)?)%$/.exec(value.trim());
  if (!match) return fallback;
  const parsed = Number.parseFloat(match[1]) / 100;
  return Number.isFinite(parsed) ? parsed : fallback;
}

interface BarBandLayout {
  /** Bar cross-section in pixels. */
  thickness: number;
  /** Cross-axis distance between the leading edges of consecutive bars in a group. */
  step: number;
  /** Leading edge of the first bar in a group, relative to the category centre. */
  firstOffset: number;
}

/**
 * Reproduce the ECharts clustered-bar layout from the gaps this renderer sets.
 *
 * ECharts fits `n` bars plus their gaps into the share of the category band left
 * over by `barCategoryGap`, with `barGap` expressed as a fraction of one bar.
 */
export function computeBarBandLayout(
  band: number,
  seriesCount: number,
  categoryGapRatio: number,
  barGapRatio: number,
): BarBandLayout | undefined {
  if (!(band > 0) || seriesCount < 1) return undefined;
  const slots = seriesCount + (seriesCount - 1) * barGapRatio;
  if (!(slots > 0)) return undefined;

  const thickness = (band * (1 - categoryGapRatio)) / slots;
  if (!(thickness > 0)) return undefined;

  const step = thickness * (1 + barGapRatio);
  return { thickness, step, firstOffset: (-thickness * slots) / 2 };
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function asSeriesArray(option: EChartsTypes.EChartsOption): SeriesRecord[] {
  const series = option.series;
  if (Array.isArray(series)) return series as SeriesRecord[];
  return series ? [series as SeriesRecord] : [];
}

function seriesPattern(series: SeriesRecord): PicturePattern | undefined {
  const itemStyle = series.itemStyle as { color?: unknown } | undefined;
  return isPicturePattern(itemStyle?.color) ? itemStyle.color : undefined;
}

function seriesData(series: SeriesRecord): unknown[] {
  return Array.isArray(series.data) ? series.data : [];
}

/** Value of a data point, which may be bare or wrapped as `{ value }`. */
function pointValue(point: unknown): unknown {
  return point && typeof point === 'object' && !Array.isArray(point)
    ? (point as { value?: unknown }).value
    : point;
}

/**
 * Attach a pattern to a data point, unless it already carries its own color
 * (a `c:dPt` override, negative-value inversion, or `varyColors`).
 */
function withPattern(point: unknown, tile: PicturePattern): unknown | undefined {
  const wrapper =
    point && typeof point === 'object' && !Array.isArray(point)
      ? (point as Record<string, unknown>)
      : undefined;
  const existingStyle = wrapper?.itemStyle as { color?: unknown } | undefined;
  if (existingStyle && existingStyle.color !== undefined) return undefined;

  return wrapper
    ? { ...wrapper, itemStyle: { ...existingStyle, color: tile } }
    : { value: point, itemStyle: { color: tile } };
}

interface NaturalSize {
  width: number;
  height: number;
}

function loadNaturalSize(url: string): Promise<NaturalSize | undefined> {
  return new Promise((resolve) => {
    const image = new Image();
    // onload rather than decode(): decoding is deferred in background tabs.
    image.onload = () =>
      resolve(
        image.naturalWidth && image.naturalHeight
          ? { width: image.naturalWidth, height: image.naturalHeight }
          : undefined,
      );
    image.onerror = () => resolve(undefined);
    image.src = url;
  });
}

// ---------------------------------------------------------------------------
// Bar and column series
// ---------------------------------------------------------------------------

interface BarContext extends BarBandLayout {
  categoryAxisKey: 'xAxisIndex' | 'yAxisIndex';
  valueAxisKey: 'xAxisIndex' | 'yAxisIndex';
  categoryOnY: boolean;
  /** Pixel of the value axis zero line, where bars start. */
  basePx: number;
  /** Bar series in draw order, for slot offsets. */
  order: SeriesRecord[];
  stacked: boolean;
}

function measureBarContext(
  chart: EChartsType,
  option: EChartsTypes.EChartsOption,
  barSeries: SeriesRecord[],
): BarContext | undefined {
  if (barSeries.length === 0) return undefined;

  const categoryOnY = (option.yAxis as { type?: string } | undefined)?.type === 'category';
  const categoryAxisKey = categoryOnY ? 'yAxisIndex' : 'xAxisIndex';
  const valueAxisKey = categoryOnY ? 'xAxisIndex' : 'yAxisIndex';

  const categoryCount = Math.max(...barSeries.map((series) => seriesData(series).length));
  // Two categories are the minimum that reveals the band width.
  if (categoryCount < 2) return undefined;

  const first = chart.convertToPixel({ [categoryAxisKey]: 0 }, 0) as number | undefined;
  const second = chart.convertToPixel({ [categoryAxisKey]: 0 }, 1) as number | undefined;
  if (typeof first !== 'number' || typeof second !== 'number') return undefined;
  const band = Math.abs(second - first);

  // Stacked series share one slot; clustered series each get their own.
  const stacked = barSeries.some((series) => typeof series.stack === 'string');
  const layout = computeBarBandLayout(
    band,
    stacked ? 1 : barSeries.length,
    parsePercentRatio(barSeries[0].barCategoryGap, 0),
    parsePercentRatio(barSeries[0].barGap, 0),
  );
  if (!layout) return undefined;

  const basePx = chart.convertToPixel({ [valueAxisKey]: 0 }, 0) as number | undefined;
  if (typeof basePx !== 'number') return undefined;

  return {
    ...layout,
    categoryAxisKey,
    valueAxisKey,
    categoryOnY,
    basePx,
    order: barSeries,
    stacked,
  };
}

function patchBarSeries(
  chart: EChartsType,
  series: SeriesRecord,
  pattern: PicturePattern,
  natural: NaturalSize,
  bar: BarContext,
): unknown[] | undefined {
  const tiled = pattern.pictureFormat === 'stack' || pattern.pictureFormat === 'stackScale';
  // Tiled: the picture runs along the bar, so its other edge spans the thickness.
  const acrossBar = bar.categoryOnY ? natural.height : natural.width;
  const tileScale = bar.thickness / acrossBar;
  const slot = bar.stacked ? 0 : bar.order.indexOf(series);
  const leadingEdge = bar.firstOffset + slot * bar.step;

  let patched = false;
  const next = seriesData(series).map((point, index) => {
    const centre = chart.convertToPixel({ [bar.categoryAxisKey]: 0 }, index) as number | undefined;
    if (typeof centre !== 'number') return point;
    const crossStart = centre + leadingEdge;

    let tile: PicturePattern;
    if (tiled) {
      tile = {
        ...pattern,
        repeat: 'repeat',
        scaleX: tileScale,
        scaleY: tileScale,
        x: bar.categoryOnY ? bar.basePx : crossStart,
        y: bar.categoryOnY ? crossStart : bar.basePx,
      };
    } else {
      // Stretched: one copy covering this bar's own rectangle.
      const raw = pointValue(point);
      if (typeof raw !== 'number' || !Number.isFinite(raw)) return point;
      const endPx = chart.convertToPixel({ [bar.valueAxisKey]: 0 }, raw) as number | undefined;
      if (typeof endPx !== 'number') return point;
      const alongSpan = Math.abs(endPx - bar.basePx);
      if (!(alongSpan > 0)) return point;

      tile = {
        ...pattern,
        repeat: 'no-repeat',
        scaleX: (bar.categoryOnY ? alongSpan : bar.thickness) / natural.width,
        scaleY: (bar.categoryOnY ? bar.thickness : alongSpan) / natural.height,
        x: bar.categoryOnY ? Math.min(bar.basePx, endPx) : crossStart,
        y: bar.categoryOnY ? crossStart : Math.min(bar.basePx, endPx),
      };
    }

    const replaced = withPattern(point, tile);
    if (!replaced) return point;
    patched = true;
    return replaced;
  });

  return patched ? next : undefined;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/**
 * Rescale and anchor every picture fill to the bar it paints.
 *
 * Resolves once the option has been patched, or immediately when the chart has
 * no picture-filled bars or cannot be measured.
 */
export async function applyBarPictureFillGeometry(
  chart: EChartsType,
  option: EChartsTypes.EChartsOption,
): Promise<void> {
  const seriesList = asSeriesArray(option);
  const barSeries = seriesList.filter((series) => series.type === 'bar');
  const patternedBars = barSeries.filter(seriesPattern);
  if (patternedBars.length === 0) return;

  const pointCount = patternedBars.reduce((sum, series) => sum + seriesData(series).length, 0);
  if (pointCount > MAX_PATTERNED_POINTS) return;

  const barContext = measureBarContext(chart, option, barSeries);
  if (!barContext) return;

  const sizes = new Map<string, NaturalSize | undefined>();
  for (const series of patternedBars) {
    const image = seriesPattern(series)!.image;
    if (!sizes.has(image)) sizes.set(image, await loadNaturalSize(image));
  }

  let patched = false;
  const nextSeries = seriesList.map((series) => {
    const pattern = seriesPattern(series);
    if (!pattern || series.type !== 'bar') return series;
    const natural = sizes.get(pattern.image);
    if (!natural) return series;

    const data = patchBarSeries(chart, series, pattern, natural, barContext);
    if (!data) return series;

    patched = true;
    return { ...series, data };
  });

  if (!patched) return;
  chart.setOption({ series: nextSeries } as EChartsTypes.EChartsOption, false);
}
