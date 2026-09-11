import { describe, it, expect } from 'vitest';
import { parseChartXml } from '../../../../src/renderer/ChartRenderer';
import { parseXml } from '../../../../src/parser/XmlParser';
import { computeWaterfallBars } from '../../../../src/renderer/chart/chartEx';
import { createMockRenderContext } from '../../helpers/mockContext';

/**
 * `cx:chartSpace` (chartex) is the Office 2016 chart format. It shares nothing
 * with `c:chartSpace` but the word "chart": data lives in a separate
 * `cx:chartData` block addressed by id, and the chart kind is an attribute
 * (`layoutId`) rather than an element name.
 */

function buildChartEx(
  options: {
    layoutId?: string;
    categories?: string[];
    values?: number[];
    subtotals?: number[];
    title?: string;
    seriesName?: string;
  } = {},
): string {
  const {
    layoutId = 'waterfall',
    categories = ['Start', 'Q1', 'Q2', 'End'],
    values = [100, 50, -30, 120],
    subtotals,
    title,
    seriesName,
  } = options;

  const cats = categories.map((c, i) => `<cx:pt idx="${i}">${c}</cx:pt>`).join('');
  const vals = values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('');

  return `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <cx:chartData>
      <cx:data id="0">
        <cx:strDim type="cat"><cx:lvl ptCount="${categories.length}">${cats}</cx:lvl></cx:strDim>
        <cx:numDim type="val"><cx:lvl ptCount="${values.length}">${vals}</cx:lvl></cx:numDim>
      </cx:data>
    </cx:chartData>
    <cx:chart>
      ${
        title
          ? `<cx:title><cx:tx><cx:rich><a:p><a:r><a:t>${title}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`
          : ''
      }
      <cx:plotArea>
        <cx:plotAreaRegion>
          <cx:series layoutId="${layoutId}" uniqueId="{TEST}">
            ${seriesName ? `<cx:tx><cx:txData><cx:v>${seriesName}</cx:v></cx:txData></cx:tx>` : ''}
            <cx:dataId val="0"/>
            ${
              subtotals
                ? `<cx:layoutPr><cx:subtotals>${subtotals
                    .map((i) => `<cx:idx val="${i}"/>`)
                    .join('')}</cx:subtotals></cx:layoutPr>`
                : ''
            }
          </cx:series>
        </cx:plotAreaRegion>
      </cx:plotArea>
    </cx:chart>
  </cx:chartSpace>`;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parse(xml: string): any {
  return parseChartXml(parseXml(xml), createMockRenderContext());
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function seriesOf(xml: string): any[] {
  const { option } = parse(xml);
  return (option.series ?? []) as any[];
}

// ---------------------------------------------------------------------------
// The waterfall running total — the substance of the format
// ---------------------------------------------------------------------------

describe('computeWaterfallBars', () => {
  it('floats each bar on the running total of the ones before it', () => {
    const bars = computeWaterfallBars([100, 50, -30], []);
    expect(bars.map((b) => b.base)).toEqual([0, 100, 120]);
    expect(bars.map((b) => b.span)).toEqual([100, 50, 30]);
  });

  it('hangs a negative bar below the running total rather than above it', () => {
    // Running total is 100; a -30 step occupies 70..100, not 100..130.
    const bars = computeWaterfallBars([100, -30], []);
    expect(bars[1]).toMatchObject({ base: 70, span: 30, direction: 'decrease' });
  });

  it('grounds a subtotal bar at zero and restarts the running total from it', () => {
    // Index 3 is declared a subtotal, so it is drawn as an absolute column.
    const bars = computeWaterfallBars([100, 50, -30, 120], [3]);
    expect(bars[3]).toMatchObject({ base: 0, span: 120, direction: 'total' });
  });

  it('continues accumulating from a subtotal, not from the pre-subtotal running total', () => {
    const bars = computeWaterfallBars([100, 50, 120, 10], [2]);
    // The subtotal resets the running total to 120, so the next bar floats there.
    expect(bars[3].base).toBe(120);
  });

  it('labels direction so increases, decreases and totals can be coloured apart', () => {
    const bars = computeWaterfallBars([100, -30, 70], [2]);
    expect(bars.map((b) => b.direction)).toEqual(['increase', 'decrease', 'total']);
  });

  it('treats a zero step as an increase of no height', () => {
    const bars = computeWaterfallBars([50, 0], []);
    expect(bars[1]).toMatchObject({ base: 50, span: 0 });
  });

  it('survives a subtotal index pointing outside the data', () => {
    expect(() => computeWaterfallBars([1, 2], [99])).not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// Dispatch and data binding
// ---------------------------------------------------------------------------

describe('chartEx', () => {
  it('is recognised rather than reported as an unsupported chart type', () => {
    const { option } = parse(buildChartEx());
    expect(option.title?.text).not.toBe('Unsupported chart type');
  });

  it('reads categories and values out of the separate chartData block', () => {
    const { option } = parse(buildChartEx({ categories: ['A', 'B', 'C'], values: [1, 2, 3] }));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xAxis = option.xAxis as any;
    expect(xAxis.data).toEqual(['A', 'B', 'C']);
  });

  it('renders a waterfall as a transparent base plus a visible span', () => {
    const series = seriesOf(buildChartEx({ values: [100, 50, -30], categories: ['A', 'B', 'C'] }));
    expect(series).toHaveLength(2);
    const [base, span] = series;
    expect(base.stack).toBe(span.stack);
    expect(base.data).toEqual([0, 100, 120]);
    expect(span.data.map((d: { value: number }) => d.value)).toEqual([100, 50, 30]);
  });

  it('hides the base bars so only the floating spans are visible', () => {
    const [base] = seriesOf(buildChartEx());
    expect(base.itemStyle?.color).toBe('transparent');
    // A stack member that is invisible must not contribute a tooltip entry.
    expect(base.tooltip?.show === false || base.silent === true).toBe(true);
  });

  it('colours increases, decreases and totals differently', () => {
    const [, span] = seriesOf(
      buildChartEx({ values: [100, -30, 70], categories: ['A', 'B', 'C'], subtotals: [2] }),
    );
    const colors = span.data.map((d: { itemStyle: { color: string } }) => d.itemStyle.color);
    expect(new Set(colors).size).toBe(3);
  });

  it('draws every funnel stage in one series colour', () => {
    // The stages are one measure narrowing, not unrelated categories; per-stage
    // palette colours would imply otherwise.
    const [funnel] = seriesOf(
      buildChartEx({
        layoutId: 'funnel',
        categories: ['Visits', 'Leads', 'Sales'],
        values: [1000, 400, 120],
      }),
    );
    const colors = funnel.data.map((d: { itemStyle?: { color?: string } }) => d.itemStyle?.color);
    expect(new Set(colors).size).toBe(1);
    expect(colors[0]).toBeTruthy();
  });

  it('takes waterfall colours from the theme palette rather than fixed values', () => {
    const [, span] = seriesOf(buildChartEx({ values: [100, -30, 70], subtotals: [2] }));
    const colors = span.data.map((d: { itemStyle: { color: string } }) => d.itemStyle.color);
    // The first three accents, in increase/decrease/total order.
    expect(colors[0]).not.toBe(colors[1]);
    expect(new Set(colors).size).toBe(3);
  });

  it('renders a funnel with its points in source order', () => {
    const series = seriesOf(
      buildChartEx({
        layoutId: 'funnel',
        categories: ['Visits', 'Leads', 'Sales'],
        values: [1000, 400, 120],
      }),
    );
    expect(series).toHaveLength(1);
    expect(series[0].type).toBe('funnel');
    // Left to itself ECharts reorders by value, which would scramble a funnel
    // whose stages do not happen to descend.
    expect(series[0].sort).toBe('none');
    expect(series[0].data.map((d: { name: string }) => d.name)).toEqual([
      'Visits',
      'Leads',
      'Sales',
    ]);
  });

  it('carries the chart title through', () => {
    const { option } = parse(buildChartEx({ title: 'Cash flow' }));
    expect(option.title?.text).toBe('Cash flow');
  });

  it('reports an unsupported layout honestly instead of drawing wrong data', () => {
    // treemap and friends are not implemented; silently drawing a waterfall
    // would misrepresent the data.
    const { option } = parse(buildChartEx({ layoutId: 'treemap' }));
    expect(option.title?.text).toContain('Unsupported');
  });

  it('handles a series whose data id matches no data block', () => {
    const xml = buildChartEx().replace('<cx:dataId val="0"/>', '<cx:dataId val="7"/>');
    expect(() => parse(xml)).not.toThrow();
  });

  it('tolerates a sparse point list', () => {
    const xml = `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <cx:chartData><cx:data id="0">
        <cx:strDim type="cat"><cx:lvl ptCount="3">
          <cx:pt idx="0">A</cx:pt><cx:pt idx="2">C</cx:pt>
        </cx:lvl></cx:strDim>
        <cx:numDim type="val"><cx:lvl ptCount="3">
          <cx:pt idx="0">10</cx:pt><cx:pt idx="2">30</cx:pt>
        </cx:lvl></cx:numDim>
      </cx:data></cx:chartData>
      <cx:chart><cx:plotArea><cx:plotAreaRegion>
        <cx:series layoutId="waterfall"><cx:dataId val="0"/></cx:series>
      </cx:plotAreaRegion></cx:plotArea></cx:chart>
    </cx:chartSpace>`;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xAxis = parse(xml).option.xAxis as any;
    // The gap keeps its slot so values stay aligned with their categories.
    expect(xAxis.data).toHaveLength(3);
    expect(xAxis.data[1]).toBe('');
  });
});
