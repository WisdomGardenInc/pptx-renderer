/**
 * Split rules for `c:ofPieChart` — the pie of pie / bar of pie chart.
 *
 * Unlike every other chart type, an of-pie does not plot its series directly.
 * The single series is divided between a primary pie and a secondary plot, and
 * the primary pie represents everything that moved out as one aggregate slice.
 * Which points move is governed by `c:splitType`, and getting that wrong is not
 * obvious from the picture — the chart still renders, just describing different
 * data — so the rule lives here on its own, away from the drawing code.
 */

import { SafeXmlNode } from '../../parser/XmlParser';

export type OfPieSplitType = 'auto' | 'pos' | 'val' | 'percent' | 'cust';

/** `c:ofPieType` — whether the secondary plot is a pie or a stacked bar. */
export type OfPieType = 'pie' | 'bar';

export interface OfPieConfig {
  ofPieType: OfPieType;
  splitType: OfPieSplitType;
  splitPos: number;
  custIndices: number[];
  /** Secondary plot size as a percentage of the primary pie. */
  secondPieSize: number;
}

interface OfPieSplit {
  /** Indices staying on the primary pie, in plot order. */
  primary: number[];
  /** Indices moved to the secondary plot, in plot order. */
  secondary: number[];
}

/**
 * Office's default split position. The spec leaves `c:splitPos` optional and
 * `auto` undefined in numeric terms; PowerPoint breaks out the last two points.
 */
const DEFAULT_SPLIT_POS = 2;

/** Below this a split leaves the primary pie with one real slice plus an aggregate — meaningless. */
const MIN_POINTS_TO_SPLIT = 3;

export function parseOfPieConfig(chartTypeNode: SafeXmlNode): OfPieConfig {
  const ofPieType = chartTypeNode.child('ofPieType').attr('val') === 'bar' ? 'bar' : 'pie';
  const rawSplit = chartTypeNode.child('splitType').attr('val');
  const splitType: OfPieSplitType =
    rawSplit === 'pos' || rawSplit === 'val' || rawSplit === 'percent' || rawSplit === 'cust'
      ? rawSplit
      : 'auto';

  const custIndices = chartTypeNode
    .child('custSplit')
    .children('secondPiePt')
    .map((pt) => pt.numAttr('val'))
    .filter((v): v is number => v !== undefined && Number.isFinite(v));

  return {
    ofPieType,
    splitType,
    splitPos: chartTypeNode.child('splitPos').numAttr('val') ?? DEFAULT_SPLIT_POS,
    custIndices,
    secondPieSize: chartTypeNode.child('secondPieSize').numAttr('val') ?? 75,
  };
}

/**
 * Decide which points belong to the secondary plot.
 *
 * Returns an empty `secondary` when the series is too small to be worth
 * splitting, which the caller renders as an ordinary pie. The primary is never
 * left empty: a split position past the end of the data would otherwise erase
 * the very pie the secondary is supposed to break out of.
 */
export function computeOfPieSplit(values: number[], config: OfPieConfig): OfPieSplit {
  const count = values.length;
  const all = Array.from({ length: count }, (_, i) => i);
  if (count < MIN_POINTS_TO_SPLIT) return { primary: all, secondary: [] };

  const total = values.reduce((sum, v) => sum + (Number.isFinite(v) ? v : 0), 0);
  const pos = Math.max(0, Math.floor(config.splitPos));

  let secondary: number[];
  switch (config.splitType) {
    case 'cust':
      secondary = config.custIndices.filter((i) => i >= 0 && i < count);
      break;
    case 'val':
      secondary = all.filter((i) => values[i] < config.splitPos);
      break;
    case 'percent':
      secondary = total > 0 ? all.filter((i) => (values[i] / total) * 100 < config.splitPos) : [];
      break;
    case 'pos':
    case 'auto':
    default:
      // Position-based splits count back from the end of the series.
      secondary = pos > 0 ? all.slice(Math.max(0, count - pos)) : [];
      break;
  }

  const secondarySet = new Set(secondary);
  let primary = all.filter((i) => !secondarySet.has(i));

  // A split that claims every point would leave nothing to break out of.
  if (primary.length === 0 && secondary.length > 0) {
    const [first, ...rest] = secondary;
    primary = [first];
    secondary = rest;
  }

  if (secondary.length === 0) return { primary: all, secondary: [] };
  return { primary, secondary: [...secondary].sort((a, b) => a - b) };
}

/**
 * A name for the aggregate slice that cannot collide with a real category —
 * the slice stands for the secondary plot, so sharing a name would make the
 * legend and tooltips ambiguous.
 */
export function aggregateSliceName(categories: string[], base = 'Other'): string {
  const taken = new Set(categories);
  if (!taken.has(base)) return base;
  for (let i = 2; ; i++) {
    const candidate = `${base} (${i})`;
    if (!taken.has(candidate)) return candidate;
  }
}
