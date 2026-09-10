import { describe, expect, it } from 'vitest';
import type * as EChartsTypes from 'echarts';
import { reserveSpaceForBottomLegend } from '../../../../src/renderer/ChartRenderer';

/** Stand-in for a laid-out element; jsdom reports zeroes for real ones. */
function elementWithRect(
  rect: { top: number; bottom: number },
  style: Partial<CSSStyleDeclaration> = {},
) {
  const el = document.createElement('div');
  Object.assign(el.style, style);
  el.getBoundingClientRect = () =>
    ({ top: rect.top, bottom: rect.bottom, height: rect.bottom - rect.top }) as DOMRect;
  return el;
}

function optionWithGridBottom(bottom: number): EChartsTypes.EChartsOption {
  return { grid: { bottom, containLabel: true } } as EChartsTypes.EChartsOption;
}

function gridBottomOf(option: EChartsTypes.EChartsOption): unknown {
  return (option.grid as { bottom?: unknown }).bottom;
}

describe('reserveSpaceForBottomLegend', () => {
  it('grows the reserve so a tall bottom legend clears the axis labels', () => {
    const option = optionWithGridBottom(35);
    // Chart spans 0..569; the legend overlay occupies 520..540.
    const chart = elementWithRect({ top: 0, bottom: 569 });
    const legend = elementWithRect({ top: 520, bottom: 540 }, { bottom: '28px' });

    reserveSpaceForBottomLegend(option, chart, legend);

    // 569 - 520 + 4px gap
    expect(gridBottomOf(option)).toBe(53);
  });

  it('keeps the existing reserve when the legend already fits', () => {
    const option = optionWithGridBottom(35);
    const chart = elementWithRect({ top: 0, bottom: 300 });
    const legend = elementWithRect({ top: 280, bottom: 292 }, { bottom: '8px' });

    reserveSpaceForBottomLegend(option, chart, legend);

    expect(gridBottomOf(option)).toBe(35);
  });

  it('never lets the legend squeeze the plot past 40% of the chart height', () => {
    const option = optionWithGridBottom(35);
    const chart = elementWithRect({ top: 0, bottom: 200 });
    const legend = elementWithRect({ top: 20, bottom: 190 }, { bottom: '10px' });

    reserveSpaceForBottomLegend(option, chart, legend);

    expect(gridBottomOf(option)).toBe(80);
  });

  it('ignores legends that are not anchored to the bottom edge', () => {
    const option = optionWithGridBottom(20);
    const chart = elementWithRect({ top: 0, bottom: 400 });
    const topLegend = elementWithRect({ top: 10, bottom: 30 }, { top: '10px' });

    reserveSpaceForBottomLegend(option, chart, topLegend);

    expect(gridBottomOf(option)).toBe(20);
  });

  it('leaves the constant alone when nothing has been laid out yet', () => {
    const option = optionWithGridBottom(35);
    const chart = elementWithRect({ top: 0, bottom: 0 });
    const legend = elementWithRect({ top: 0, bottom: 0 }, { bottom: '5px' });

    reserveSpaceForBottomLegend(option, chart, legend);
    reserveSpaceForBottomLegend(option, chart, null);

    expect(gridBottomOf(option)).toBe(35);
  });

  it('does not shrink a reserve that the option already set larger than the cap', () => {
    const option = optionWithGridBottom(120);
    const chart = elementWithRect({ top: 0, bottom: 200 });
    const legend = elementWithRect({ top: 180, bottom: 195 }, { bottom: '5px' });

    reserveSpaceForBottomLegend(option, chart, legend);

    expect(gridBottomOf(option)).toBe(120);
  });
});
