import { describe, it, expect } from 'vitest';
import { parseChartXml } from '../../../../src/renderer/ChartRenderer';
import { parseXml } from '../../../../src/parser/XmlParser';
import { createMockRenderContext } from '../../helpers/mockContext';

/**
 * `c:ofPieChart` is a pie of pie / bar of pie chart. Unlike every other chart
 * type it does not simply plot its series: the one series is *split* between a
 * primary pie and a secondary plot, and the primary pie shows the split-off
 * points collapsed into a single aggregate slice. The split rule is what these
 * tests pin down — drawing two pies is the easy half.
 */

interface Ser {
  categories: string[];
  values: number[];
}

function buildOfPie(
  options: {
    ofPieType?: string;
    splitType?: string;
    splitPos?: number;
    custSplit?: number[];
    secondPieSize?: number;
    ser?: Ser;
  } = {},
): string {
  const {
    ofPieType = 'pie',
    splitType,
    splitPos,
    custSplit,
    secondPieSize,
    ser = {
      categories: ['A', 'B', 'C', 'D', 'E'],
      values: [40, 30, 10, 6, 4],
    },
  } = options;

  const cats = ser.categories
    .map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`)
    .join('');
  const vals = ser.values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');

  return `
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        <c:plotArea>
          <c:ofPieChart>
            <c:ofPieType val="${ofPieType}"/>
            <c:varyColors val="1"/>
            <c:ser>
              <c:idx val="0"/><c:order val="0"/>
              <c:tx><c:strRef><c:strCache><c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Share</c:v></c:pt>
              </c:strCache></c:strRef></c:tx>
              <c:cat><c:strRef><c:strCache><c:ptCount val="${ser.categories.length}"/>
                ${cats}
              </c:strCache></c:strRef></c:cat>
              <c:val><c:numRef><c:numCache><c:ptCount val="${ser.values.length}"/>
                ${vals}
              </c:numCache></c:numRef></c:val>
            </c:ser>
            ${splitType ? `<c:splitType val="${splitType}"/>` : ''}
            ${splitPos !== undefined ? `<c:splitPos val="${splitPos}"/>` : ''}
            ${
              custSplit
                ? `<c:custSplit>${custSplit
                    .map((i) => `<c:secondPiePt val="${i}"/>`)
                    .join('')}</c:custSplit>`
                : ''
            }
            ${secondPieSize !== undefined ? `<c:secondPieSize val="${secondPieSize}"/>` : ''}
          </c:ofPieChart>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>`;
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

describe('ofPieChart', () => {
  it('is recognised rather than reported as an unsupported chart type', () => {
    const { option } = parse(buildOfPie());
    expect(option.title?.text).not.toBe('Unsupported chart type');
    expect(option.series).toBeDefined();
  });

  it('splits the series into a primary and a secondary plot', () => {
    const series = seriesOf(buildOfPie());
    expect(series).toHaveLength(2);
    expect(series[0].type).toBe('pie');
    expect(series[1].type).toBe('pie');
  });

  it('defaults to moving the last two points into the secondary pie', () => {
    // Office's default is splitType "auto", which behaves as position-based with
    // a split position of 2.
    const series = seriesOf(buildOfPie());
    const secondary = series[1].data.map((d: { name: string }) => d.name);
    expect(secondary).toEqual(['D', 'E']);
  });

  it('collapses the split-off points into one aggregate slice on the primary pie', () => {
    const series = seriesOf(buildOfPie());
    const primary = series[0].data;
    // A, B, C stay; D+E (6+4) become a single 10-valued slice.
    expect(primary.map((d: { name: string }) => d.name).slice(0, 3)).toEqual(['A', 'B', 'C']);
    expect(primary).toHaveLength(4);
    expect(primary[3].value).toBe(10);
  });

  it('honours an explicit position split', () => {
    const series = seriesOf(buildOfPie({ splitType: 'pos', splitPos: 3 }));
    expect(series[1].data.map((d: { name: string }) => d.name)).toEqual(['C', 'D', 'E']);
    expect(series[0].data).toHaveLength(3); // A, B, aggregate
    expect(series[0].data[2].value).toBe(20);
  });

  // These two share a series and a threshold and differ only in splitType, so
  // they pin the distinction between an absolute and a relative cutoff. The
  // total is 200, which keeps a value of 10 and 10% of the total far apart.
  const scaledSer = {
    categories: ['A', 'B', 'C', 'D', 'E'],
    values: [100, 60, 20, 12, 8],
  };

  it('splits by value when splitType is val', () => {
    const series = seriesOf(buildOfPie({ ser: scaledSer, splitType: 'val', splitPos: 10 }));
    // Only E is below an absolute value of 10.
    expect(series[1].data.map((d: { name: string }) => d.name)).toEqual(['E']);
  });

  it('splits by percentage when splitType is percent', () => {
    const series = seriesOf(buildOfPie({ ser: scaledSer, splitType: 'percent', splitPos: 10 }));
    // D is 6% and E is 4% of the 200 total, so both fall below 10%.
    expect(series[1].data.map((d: { name: string }) => d.name)).toEqual(['D', 'E']);
  });

  it('uses the explicit point list when splitType is cust', () => {
    const series = seriesOf(buildOfPie({ splitType: 'cust', custSplit: [0, 4] }));
    expect(series[1].data.map((d: { name: string }) => d.name)).toEqual(['A', 'E']);
    expect(series[0].data.map((d: { name: string }) => d.name)).toEqual([
      'B',
      'C',
      'D',
      'Other',
    ]);
  });

  it('renders the secondary plot as a bar when ofPieType is bar', () => {
    const series = seriesOf(buildOfPie({ ofPieType: 'bar' }));
    expect(series[0].type).toBe('pie');
    expect(series[1].type).toBe('bar');
    // A bar of pie stacks its points in one column.
    expect(series[1].stack).toBeTruthy();
  });

  it('sizes the bar-of-pie column to fill its plot rather than the pie\'s range', () => {
    // The column expands the aggregate slice, so its stack total is the axis
    // maximum. Left to the generic axis pass it would be scaled against the
    // pie's own values and collapse to a sliver.
    const { option } = parse(buildOfPie({ ofPieType: 'bar', splitType: 'pos', splitPos: 3 }));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const yAxis = option.yAxis as any;
    expect(yAxis.min).toBe(0);
    expect(yAxis.max).toBe(20); // C + D + E = 10 + 6 + 4
  });

  it('places the two plots side by side so they do not overlap', () => {
    const series = seriesOf(buildOfPie());
    const [primary, secondary] = series;
    const centreX = (s: { center: [string, string] }) => parseFloat(s.center[0]);
    expect(centreX(primary)).toBeLessThan(centreX(secondary));
  });

  it('scales the secondary pie by secondPieSize', () => {
    const bigger = seriesOf(buildOfPie({ secondPieSize: 100 }));
    const smaller = seriesOf(buildOfPie({ secondPieSize: 50 }));
    const radius = (s: { radius: string | string[] }) =>
      parseFloat(Array.isArray(s.radius) ? s.radius[s.radius.length - 1] : s.radius);
    expect(radius(bigger[1])).toBeGreaterThan(radius(smaller[1]));
  });

  it('keeps every point when the split would empty the primary pie', () => {
    // A split position at or beyond the point count would leave nothing behind;
    // Office still draws a primary pie, so the split is clamped.
    const series = seriesOf(buildOfPie({ splitType: 'pos', splitPos: 99 }));
    expect(series[0].data.length).toBeGreaterThan(0);
  });

  it('falls back to a plain pie when the series has too few points to split', () => {
    const series = seriesOf(
      buildOfPie({ ser: { categories: ['A', 'B'], values: [60, 40] }, splitType: 'pos', splitPos: 2 }),
    );
    // Nothing meaningful to break out, so one pie carrying both points.
    expect(series).toHaveLength(1);
    expect(series[0].data).toHaveLength(2);
  });

  it('never colours the aggregate slice like one already on the primary pie', () => {
    // A six-point series against a six-entry theme palette wraps an index-based
    // colour straight back onto the first slice, making the two indistinguishable.
    const ser = {
      categories: ['A', 'B', 'C', 'D', 'E', 'F'],
      values: [60, 20, 8, 6, 4, 2],
    };
    const series = seriesOf(buildOfPie({ ser, splitType: 'pos', splitPos: 3 }));
    const primary = series[0].data as Array<{ itemStyle?: { color?: string } }>;
    const colors = primary.map((d) => d.itemStyle?.color).filter(Boolean);
    expect(new Set(colors).size).toBe(colors.length);
  });

  it('gives the aggregate slice and its exploded points a shared identity', () => {
    // The aggregate slice stands for the secondary plot, so it must not reuse a
    // category name that already exists on the primary pie.
    const series = seriesOf(buildOfPie());
    const primaryNames = series[0].data.map((d: { name: string }) => d.name);
    const secondaryNames = series[1].data.map((d: { name: string }) => d.name);
    expect(new Set(primaryNames).size).toBe(primaryNames.length);
    for (const name of secondaryNames) {
      expect(primaryNames).not.toContain(name);
    }
  });
});
