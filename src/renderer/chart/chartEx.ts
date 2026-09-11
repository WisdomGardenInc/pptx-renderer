/**
 * chartex (`cx:chartSpace`) — the chart format Office 2016 introduced for
 * waterfall, funnel, treemap, sunburst, box & whisker, histogram and Pareto.
 *
 * It shares nothing with the original `c:chartSpace` but the word "chart":
 *
 *  - Data lives in a separate `cx:chartData` block of dimensions addressed by
 *    id, instead of inline under each series.
 *  - The chart kind is an attribute — `cx:series/@layoutId` — rather than the
 *    element name, so one parser handles every kind.
 *  - Categories can be multi-level (treemap, sunburst) and points are sparse,
 *    carrying their own `idx`.
 *
 * Only the layouts listed in `SUPPORTED_LAYOUTS` are drawn. Anything else is
 * reported as unsupported rather than approximated: these charts differ in what
 * the numbers *mean*, so drawing a treemap's data as a waterfall would be a
 * confident lie rather than a rough likeness.
 */

import type * as EChartsTypes from 'echarts';
import { SafeXmlNode } from '../../parser/XmlParser';

/** Chart kinds this module can draw, keyed by `cx:series/@layoutId`. */
const SUPPORTED_LAYOUTS = new Set(['waterfall', 'funnel']);

/**
 * Office's own waterfall colours, used when the deck supplies no theme palette.
 * Increases, decreases and totals have to read apart at a glance — that
 * distinction is the point of the chart type.
 */
const WATERFALL_FALLBACK_COLORS: Record<WaterfallDirection, string> = {
  increase: '#4472C4',
  decrease: '#ED7D31',
  total: '#A5A5A5',
};

/**
 * Waterfall bars take the first three theme accents, as PowerPoint does, so the
 * chart matches the rest of the deck rather than a palette of our own choosing.
 */
function waterfallColors(palette?: string[]): Record<WaterfallDirection, string> {
  if (!palette || palette.length < 3) return WATERFALL_FALLBACK_COLORS;
  return { increase: palette[0], decrease: palette[1], total: palette[2] };
}

type WaterfallDirection = 'increase' | 'decrease' | 'total';

interface WaterfallBar {
  /** Height of the invisible spacer holding the bar up. */
  base: number;
  /** Visible height of the bar. */
  span: number;
  direction: WaterfallDirection;
}

/**
 * Turn a series of steps into floating bars.
 *
 * Each step is drawn relative to the total of everything before it, so a bar's
 * position carries as much meaning as its height. A negative step hangs *below*
 * the running total — it occupies the band the total is about to vacate — and a
 * subtotal is an absolute column measured from zero that also restarts the
 * running total, since the figures after it accumulate from that subtotal.
 */
export function computeWaterfallBars(values: number[], subtotalIndices: number[]): WaterfallBar[] {
  const subtotals = new Set(subtotalIndices);
  const bars: WaterfallBar[] = [];
  let running = 0;

  for (let i = 0; i < values.length; i++) {
    const value = Number.isFinite(values[i]) ? values[i] : 0;

    if (subtotals.has(i)) {
      bars.push({ base: 0, span: value, direction: 'total' });
      running = value;
      continue;
    }

    if (value < 0) {
      running += value;
      bars.push({ base: running, span: -value, direction: 'decrease' });
    } else {
      bars.push({ base: running, span: value, direction: 'increase' });
      running += value;
    }
  }

  return bars;
}

/** True when this part is a chartex chart rather than a classic one. */
export function isChartExSpace(chartXml: SafeXmlNode): boolean {
  return chartXml.child('chartData').exists();
}

interface ChartExDimension {
  categories: string[];
  values: number[];
  formatCode?: string;
}

/**
 * Read one `cx:data` block.
 *
 * Points are sparse and carry their own index, so slots without a point keep
 * their place — dropping them would slide every later value onto the wrong
 * category.
 */
function parseDataBlock(dataNode: SafeXmlNode): ChartExDimension {
  const readLevel = (dim: SafeXmlNode): { texts: string[]; count: number } => {
    const lvl = dim.child('lvl');
    const count = lvl.numAttr('ptCount') ?? 0;
    const texts: string[] = new Array(count).fill('');
    for (const pt of lvl.children('pt')) {
      const idx = pt.numAttr('idx');
      if (idx === undefined || idx < 0 || idx >= count) continue;
      texts[idx] = pt.text() ?? '';
    }
    return { texts, count };
  };

  const catDim =
    dataNode.children('strDim').find((d) => d.attr('type') === 'cat') ?? dataNode.child('strDim');
  const valDim =
    dataNode.children('numDim').find((d) => d.attr('type') === 'val') ?? dataNode.child('numDim');

  const cats = catDim.exists() ? readLevel(catDim) : { texts: [] as string[], count: 0 };
  const vals = valDim.exists() ? readLevel(valDim) : { texts: [] as string[], count: 0 };

  const length = Math.max(cats.count, vals.count);
  const categories = new Array(length).fill('').map((_, i) => cats.texts[i] ?? '');
  const values = new Array(length).fill(0).map((_, i) => {
    const raw = vals.texts[i];
    const n = raw === '' || raw === undefined ? NaN : Number(raw);
    return Number.isFinite(n) ? n : 0;
  });

  return {
    categories,
    values,
    formatCode: valDim.child('lvl').attr('formatCode'),
  };
}

/** Locate the `cx:data` block a series points at through its `cx:dataId`. */
function findSeriesData(chartXml: SafeXmlNode, series: SafeXmlNode): ChartExDimension | undefined {
  const dataId = series.child('dataId').attr('val');
  const blocks = chartXml.child('chartData').children('data');
  const matched = dataId !== undefined ? blocks.find((b) => b.attr('id') === dataId) : blocks[0];
  const block = matched ?? blocks[0];
  return block?.exists() ? parseDataBlock(block) : undefined;
}

function parseChartExTitle(chartXml: SafeXmlNode): string | undefined {
  const runs = chartXml
    .child('chart')
    .child('title')
    .child('tx')
    .child('rich')
    .children('p')
    .flatMap((p) => p.children('r').map((r) => r.child('t').text() ?? ''));
  const text = runs.join('').trim();
  return text.length > 0 ? text : undefined;
}

/**
 * Convert a chartex part into an ECharts option.
 *
 * Returns undefined when the part carries no series at all; an unsupported
 * layout yields a titled placeholder so the caller can tell the two apart.
 */
export function parseChartExOption(
  chartXml: SafeXmlNode,
  chartPalette?: string[],
): EChartsTypes.EChartsOption | undefined {
  const series = chartXml.child('chart').child('plotArea').child('plotAreaRegion').child('series');
  if (!series.exists()) return undefined;

  const layoutId = series.attr('layoutId') ?? '';
  const title = parseChartExTitle(chartXml);
  const titleOption = title ? { text: title, left: 'center' as const } : undefined;

  if (!SUPPORTED_LAYOUTS.has(layoutId)) {
    return {
      title: {
        text: `Unsupported chart layout: ${layoutId || 'unknown'}`,
        left: 'center',
        textStyle: { fontSize: 12 },
      },
    };
  }

  const data = findSeriesData(chartXml, series);
  if (!data || data.values.length === 0) {
    return { ...(titleOption ? { title: titleOption } : {}) };
  }

  const seriesName = series.child('tx').child('txData').child('v').text() ?? '';

  return layoutId === 'funnel'
    ? buildFunnelOption(data, seriesName, titleOption, chartPalette)
    : buildWaterfallOption(series, data, seriesName, titleOption, chartPalette);
}

function buildWaterfallOption(
  series: SafeXmlNode,
  data: ChartExDimension,
  seriesName: string,
  titleOption: { text: string; left: 'center' } | undefined,
  chartPalette: string[] | undefined,
): EChartsTypes.EChartsOption {
  const subtotals = series
    .child('layoutPr')
    .child('subtotals')
    .children('idx')
    .map((idx) => idx.numAttr('val'))
    .filter((v): v is number => v !== undefined);

  const bars = computeWaterfallBars(data.values, subtotals);
  const colors = waterfallColors(chartPalette);

  return {
    ...(titleOption ? { title: titleOption } : {}),
    tooltip: { trigger: 'axis' as const },
    xAxis: { type: 'category' as const, data: data.categories },
    yAxis: { type: 'value' as const },
    series: [
      {
        // The spacer that floats each bar; invisible and inert so it neither
        // shows up in the picture nor in tooltips.
        type: 'bar' as const,
        name: 'waterfall-base',
        stack: 'waterfall',
        silent: true,
        itemStyle: { color: 'transparent' },
        emphasis: { itemStyle: { color: 'transparent' } },
        tooltip: { show: false },
        data: bars.map((b) => b.base),
      },
      {
        type: 'bar' as const,
        name: seriesName || 'Series',
        stack: 'waterfall',
        data: bars.map((b) => ({
          value: b.span,
          itemStyle: { color: colors[b.direction] },
        })),
      },
    ],
  };
}

function buildFunnelOption(
  data: ChartExDimension,
  seriesName: string,
  titleOption: { text: string; left: 'center' } | undefined,
  chartPalette: string[] | undefined,
): EChartsTypes.EChartsOption {
  // A funnel's stages are one measure narrowing, not separate categories, so
  // PowerPoint draws them in a single series colour. Left alone ECharts would
  // give each stage its own palette entry and imply they are unrelated.
  const stageColor = chartPalette?.[0];
  return {
    ...(titleOption ? { title: titleOption } : {}),
    tooltip: { trigger: 'item' as const },
    series: [
      {
        type: 'funnel' as const,
        name: seriesName || 'Series',
        // Stages are a sequence, not a ranking: ECharts would otherwise reorder
        // them by value and misstate a funnel that does not descend.
        sort: 'none' as const,
        left: '10%',
        width: '80%',
        label: { show: true, position: 'inside' as const },
        ...(stageColor ? { itemStyle: { color: stageColor } } : {}),
        data: data.categories.map((name, i) => ({
          name: name || `Item ${i + 1}`,
          value: data.values[i] ?? 0,
          ...(stageColor ? { itemStyle: { color: stageColor } } : {}),
        })),
      },
    ],
  };
}
