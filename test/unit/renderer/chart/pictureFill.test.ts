import { describe, expect, it, vi } from 'vitest';
import type * as EChartsTypes from 'echarts';
import type { EChartsType } from 'echarts/core';
import {
  applyBarPictureFillGeometry,
  computeBarBandLayout,
  isPicturePattern,
  parsePercentRatio,
  type PicturePattern,
} from '../../../../src/renderer/chart/pictureFill';

const IMAGE = 'blob:strip';

/** Stub the two chart APIs the geometry pass uses, over a fixed layout. */
function fakeChart(options: {
  categoryOnY: boolean;
  bandPx: number;
  firstCentre: number;
  basePx: number;
  /** Pixels per unit along the value axis. */
  valueScale: number;
}) {
  const { categoryOnY, bandPx, firstCentre, basePx, valueScale } = options;
  const categoryKey = categoryOnY ? 'yAxisIndex' : 'xAxisIndex';
  const convertToPixel = vi.fn((finder: Record<string, number>, value: number) => {
    if (categoryKey in finder) return firstCentre + value * bandPx;
    return basePx + value * valueScale;
  });
  const setOption = vi.fn();
  return { convertToPixel, setOption, isDisposed: () => false } as unknown as EChartsType & {
    setOption: ReturnType<typeof vi.fn>;
  };
}

function barOption(
  categoryOnY: boolean,
  series: Record<string, unknown>[],
): EChartsTypes.EChartsOption {
  return {
    xAxis: { type: categoryOnY ? 'value' : 'category' },
    yAxis: { type: categoryOnY ? 'category' : 'value' },
    series,
  } as EChartsTypes.EChartsOption;
}

function pattern(format: PicturePattern['pictureFormat']): PicturePattern {
  return { image: IMAGE, repeat: 'repeat', pictureFormat: format };
}

/** Stub Image so naturalWidth/Height resolve without a real decode. */
function stubImage(width: number, height: number) {
  class FakeImage {
    naturalWidth = width;
    naturalHeight = height;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  vi.stubGlobal('Image', FakeImage);
}

function patchedSeries(chart: { setOption: ReturnType<typeof vi.fn> }) {
  return (chart.setOption.mock.calls[0][0] as { series: Record<string, unknown>[] }).series;
}

function tileOf(series: Record<string, unknown>[], seriesIndex: number, pointIndex: number) {
  const data = series[seriesIndex].data as Record<string, unknown>[];
  return (data[pointIndex].itemStyle as { color: PicturePattern }).color;
}

describe('picture fill helpers', () => {
  it('recognises pattern fills and ignores plain colors', () => {
    expect(isPicturePattern({ image: 'blob:x', repeat: 'repeat' })).toBe(true);
    expect(isPicturePattern({ image: 'blob:x', repeat: 'no-repeat' })).toBe(true);
    expect(isPicturePattern('#AB0515')).toBe(false);
    expect(isPicturePattern({ colorStops: [] })).toBe(false);
    expect(isPicturePattern(undefined)).toBe(false);
  });

  it('reads ECharts percentage ratios', () => {
    expect(parsePercentRatio('13%', 0)).toBeCloseTo(0.13, 10);
    expect(parsePercentRatio('-5%', 0)).toBeCloseTo(-0.05, 10);
    expect(parsePercentRatio(0.4, 0)).toBe(0.4);
    expect(parsePercentRatio('nonsense', 0.2)).toBe(0.2);
    expect(parsePercentRatio(undefined, 0.2)).toBe(0.2);
  });

  it('reproduces the ECharts clustered-bar geometry', () => {
    // Matches a measured render: band 118.64, two bars, 13% category gap, -5% bar gap.
    const layout = computeBarBandLayout(118.64, 2, 0.13, -0.05)!;
    expect(layout.thickness).toBeCloseTo(52.93, 1);
    expect(layout.step).toBeCloseTo(50.28, 1);
    expect(layout.firstOffset).toBeCloseTo(-51.6, 1);
  });

  it('rejects degenerate bands', () => {
    expect(computeBarBandLayout(0, 2, 0.1, 0)).toBeUndefined();
    expect(computeBarBandLayout(100, 0, 0.1, 0)).toBeUndefined();
    expect(computeBarBandLayout(100, 1, 1, 0)).toBeUndefined();
  });
});

describe('applyBarPictureFillGeometry', () => {
  it('scales a stacked fill to the bar thickness and anchors it at the value base', async () => {
    stubImage(1032, 86); // wide strip, as decks pair with horizontal bars
    const chart = fakeChart({
      categoryOnY: true,
      bandPx: 100,
      firstCentre: 50,
      basePx: 200,
      valueScale: 10,
    });
    const option = barOption(true, [
      {
        type: 'bar',
        data: [1, 2],
        itemStyle: { color: pattern('stack') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
    ]);

    await applyBarPictureFillGeometry(chart, option);

    const tile = tileOf(patchedSeries(chart), 0, 0);
    expect(tile.repeat).toBe('repeat');
    // One series, no gaps → thickness is the whole band; scale = 100 / image height.
    expect(tile.scaleX).toBeCloseTo(100 / 86, 6);
    expect(tile.scaleY).toBeCloseTo(100 / 86, 6);
    // Horizontal bars tile from the value base, and each bar sits on its own row.
    expect(tile.x).toBe(200);
    expect(tile.y).toBeCloseTo(0, 6);
    expect(tileOf(patchedSeries(chart), 0, 1).y).toBeCloseTo(100, 6);
  });

  it('scales a stacked fill by image width for columns', async () => {
    stubImage(118, 874); // tall strip, as decks pair with columns
    const chart = fakeChart({
      categoryOnY: false,
      bandPx: 100,
      firstCentre: 50,
      basePx: 400,
      valueScale: -10,
    });
    const option = barOption(false, [
      {
        type: 'bar',
        data: [1, 2],
        itemStyle: { color: pattern('stack') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
    ]);

    await applyBarPictureFillGeometry(chart, option);

    const tile = tileOf(patchedSeries(chart), 0, 0);
    expect(tile.scaleX).toBeCloseTo(100 / 118, 6);
    expect(tile.scaleY).toBeCloseTo(100 / 118, 6);
    expect(tile.y).toBe(400);
  });

  it('stretches one copy over each bar when no pictureFormat is given', async () => {
    stubImage(508, 509);
    const chart = fakeChart({
      categoryOnY: false,
      bandPx: 100,
      firstCentre: 50,
      basePx: 400,
      valueScale: -4, // value 25 → 100px tall
    });
    const option = barOption(false, [
      {
        type: 'bar',
        data: [25, 10],
        itemStyle: { color: pattern('stretch') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
    ]);

    await applyBarPictureFillGeometry(chart, option);

    const tile = tileOf(patchedSeries(chart), 0, 0);
    expect(tile.repeat).toBe('no-repeat');
    expect(tile.scaleX).toBeCloseTo(100 / 508, 6); // column width
    expect(tile.scaleY).toBeCloseTo(100 / 509, 6); // bar length
    expect(tile.x).toBeCloseTo(0, 6);
    expect(tile.y).toBeCloseTo(300, 6); // top of a bar rising from y=400
  });

  it('offsets clustered series so each bar gets its own tile origin', async () => {
    stubImage(1032, 86);
    const chart = fakeChart({
      categoryOnY: true,
      bandPx: 100,
      firstCentre: 50,
      basePx: 0,
      valueScale: 10,
    });
    const option = barOption(true, [
      {
        type: 'bar',
        data: [1, 2],
        itemStyle: { color: pattern('stack') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
      {
        type: 'bar',
        data: [1, 2],
        itemStyle: { color: pattern('stack') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
    ]);

    await applyBarPictureFillGeometry(chart, option);

    const series = patchedSeries(chart);
    expect(tileOf(series, 0, 0).y).toBeCloseTo(0, 6);
    expect(tileOf(series, 1, 0).y).toBeCloseTo(50, 6); // second slot, half a band down
  });

  it('leaves data points that carry their own color alone', async () => {
    stubImage(1032, 86);
    const chart = fakeChart({
      categoryOnY: true,
      bandPx: 100,
      firstCentre: 50,
      basePx: 0,
      valueScale: 10,
    });
    const option = barOption(true, [
      {
        type: 'bar',
        data: [{ value: 1, itemStyle: { color: '#FFFFFF' } }, 2],
        itemStyle: { color: pattern('stack') },
        barCategoryGap: '0%',
        barGap: '0%',
      },
    ]);

    await applyBarPictureFillGeometry(chart, option);

    const data = patchedSeries(chart)[0].data as Record<string, unknown>[];
    expect((data[0].itemStyle as { color: unknown }).color).toBe('#FFFFFF');
    expect(isPicturePattern((data[1].itemStyle as { color: unknown }).color)).toBe(true);
  });

  it('does nothing without picture-filled bars', async () => {
    stubImage(10, 10);
    const chart = fakeChart({
      categoryOnY: true,
      bandPx: 100,
      firstCentre: 50,
      basePx: 0,
      valueScale: 10,
    });

    await applyBarPictureFillGeometry(
      chart,
      barOption(true, [{ type: 'bar', data: [1, 2], itemStyle: { color: '#AB0515' } }]),
    );

    expect(chart.setOption).not.toHaveBeenCalled();
  });

  it('keeps the plain tiling when the image cannot be measured', async () => {
    stubImage(0, 0);
    const chart = fakeChart({
      categoryOnY: true,
      bandPx: 100,
      firstCentre: 50,
      basePx: 0,
      valueScale: 10,
    });

    await applyBarPictureFillGeometry(
      chart,
      barOption(true, [
        { type: 'bar', data: [1, 2], itemStyle: { color: pattern('stack') }, barCategoryGap: '0%' },
      ]),
    );

    expect(chart.setOption).not.toHaveBeenCalled();
  });
});
