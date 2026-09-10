/**
 * Geometry for OOXML `stack` picture fills on bar and column series.
 *
 * PowerPoint lays a stacked picture along the bar: the picture is scaled so the
 * edge across the bar matches the bar's thickness and anchored at the bar's base,
 * so a strip of motifs reads as whole motifs the size of the bar. Which edge that
 * is depends on the bar direction — decks pair horizontal bars with a wide strip
 * and columns with a tall one. An ECharts pattern is shape-independent:
 * it tiles at the image's natural pixel size from the canvas origin, which both
 * crops the artwork and lands at a different phase on every bar.
 *
 * Both corrections need the laid-out bar rectangles, so they are applied after
 * the chart's first render, as a per-data-point pattern transform.
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

/** Charts past this many bars keep the plain tiling; per-point patterns would not pay off. */
const MAX_PATTERNED_BARS = 2000;

export function isPicturePattern(value: unknown): value is PicturePattern {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as PicturePattern;
  return (
    typeof candidate.image === 'string' &&
    (candidate.repeat === 'repeat' || candidate.repeat === 'no-repeat')
  );
}

/** Numeric value of a data point, which may be a bare number or `{ value }`. */
function pointValue(point: unknown): number | undefined {
  const raw =
    point && typeof point === 'object' ? (point as { value?: unknown }).value : (point as unknown);
  return typeof raw === 'number' && Number.isFinite(raw) ? raw : undefined;
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
  const groupSpan = thickness * slots;
  return { thickness, step, firstOffset: -groupSpan / 2 };
}

function asSeriesArray(option: EChartsTypes.EChartsOption): SeriesRecord[] {
  const series = option.series;
  if (Array.isArray(series)) return series as SeriesRecord[];
  return series ? [series as SeriesRecord] : [];
}

function seriesPattern(series: SeriesRecord): PicturePattern | undefined {
  const itemStyle = series.itemStyle as { color?: unknown } | undefined;
  return isPicturePattern(itemStyle?.color) ? itemStyle.color : undefined;
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

/** Category-axis pixel positions, or undefined when the chart cannot be probed. */
function categoryBandPx(
  chart: EChartsType,
  axisKey: 'xAxisIndex' | 'yAxisIndex',
  categoryCount: number,
): number | undefined {
  if (categoryCount < 2) return undefined;
  const first = chart.convertToPixel({ [axisKey]: 0 }, 0) as number | undefined;
  const second = chart.convertToPixel({ [axisKey]: 0 }, 1) as number | undefined;
  if (typeof first !== 'number' || typeof second !== 'number') return undefined;
  const band = Math.abs(second - first);
  return band > 0 ? band : undefined;
}

/**
 * Rescale and anchor every `stack` picture fill to the bars it paints.
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
  if (barSeries.length === 0) return;
  if (!barSeries.some(seriesPattern)) return;

  const categoryOnY = (option.yAxis as { type?: string } | undefined)?.type === 'category';
  const categoryAxisKey = categoryOnY ? 'yAxisIndex' : 'xAxisIndex';
  const valueAxisKey = categoryOnY ? 'xAxisIndex' : 'yAxisIndex';

  const categoryCount = Math.max(
    ...barSeries.map((series) => (Array.isArray(series.data) ? series.data.length : 0)),
  );
  if (categoryCount * barSeries.length > MAX_PATTERNED_BARS) return;

  const band = categoryBandPx(chart, categoryAxisKey, categoryCount);
  if (band === undefined) return;

  // Stacked series share one slot; clustered series each get their own.
  const stacked = barSeries.some((series) => typeof series.stack === 'string');
  const slotCount = stacked ? 1 : barSeries.length;
  const layout = computeBarBandLayout(
    band,
    slotCount,
    parsePercentRatio(barSeries[0].barCategoryGap, 0),
    parsePercentRatio(barSeries[0].barGap, 0),
  );
  if (!layout) return;

  // Bars grow from the value axis zero line; that edge anchors the tiling.
  const basePx = chart.convertToPixel({ [valueAxisKey]: 0 }, 0) as number | undefined;
  if (typeof basePx !== 'number') return;

  const sizes = new Map<string, NaturalSize | undefined>();
  for (const series of barSeries) {
    const pattern = seriesPattern(series);
    if (pattern && !sizes.has(pattern.image)) {
      sizes.set(pattern.image, await loadNaturalSize(pattern.image));
    }
  }

  let patched = false;
  const nextSeries = seriesList.map((series) => {
    const pattern = seriesPattern(series);
    if (!pattern || series.type !== 'bar') return series;

    const natural = sizes.get(pattern.image);
    if (!natural) return series;

    const tiled = pattern.pictureFormat === 'stack' || pattern.pictureFormat === 'stackScale';
    // Tiled: the picture runs along the bar, so its other edge spans the thickness.
    const acrossBar = categoryOnY ? natural.height : natural.width;
    const tileScale = layout.thickness / acrossBar;
    const slot = stacked ? 0 : barSeries.indexOf(series);
    const leadingEdge = layout.firstOffset + slot * layout.step;
    const data = Array.isArray(series.data) ? series.data : [];

    const nextData = data.map((point, index) => {
      const centre = chart.convertToPixel({ [categoryAxisKey]: 0 }, index) as number | undefined;
      if (typeof centre !== 'number') return point;

      const crossStart = centre + leadingEdge;
      let tile: PicturePattern;

      if (tiled) {
        tile = {
          ...pattern,
          repeat: 'repeat',
          scaleX: tileScale,
          scaleY: tileScale,
          x: categoryOnY ? basePx : crossStart,
          y: categoryOnY ? crossStart : basePx,
        };
      } else {
        // Stretched: one copy covering this bar's own rectangle.
        const value = pointValue(point);
        if (value === undefined) return point;
        const endPx = chart.convertToPixel({ [valueAxisKey]: 0 }, value) as number | undefined;
        if (typeof endPx !== 'number') return point;
        const alongStart = Math.min(basePx, endPx);
        const alongSpan = Math.abs(endPx - basePx);
        if (!(alongSpan > 0)) return point;

        tile = {
          ...pattern,
          repeat: 'no-repeat',
          scaleX: (categoryOnY ? alongSpan : layout.thickness) / natural.width,
          scaleY: (categoryOnY ? layout.thickness : alongSpan) / natural.height,
          x: categoryOnY ? alongStart : crossStart,
          y: categoryOnY ? crossStart : alongStart,
        };
      }

      const existing =
        point && typeof point === 'object' ? (point as Record<string, unknown>) : undefined;
      const existingStyle = existing?.itemStyle as { color?: unknown } | undefined;
      // A data point with its own explicit color (dPt override, negative-value
      // inversion, varyColors) keeps it.
      if (existingStyle && existingStyle.color !== undefined) return point;

      patched = true;
      return existing
        ? { ...existing, itemStyle: { ...existingStyle, color: tile } }
        : { value: point, itemStyle: { color: tile } };
    });

    return { ...series, data: nextData };
  });

  if (!patched) return;
  chart.setOption({ series: nextSeries } as EChartsTypes.EChartsOption, false);
}
