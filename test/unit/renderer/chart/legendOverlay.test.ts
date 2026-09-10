import { describe, expect, it } from 'vitest';
import {
  buildCustomLegendOverlay,
  createLegendIcon,
} from '../../../../src/renderer/chart/legendOverlay';

describe('chart legend overlay helpers', () => {
  it('creates SVG legend icons for line paths with markers', () => {
    const icon = createLegendIcon('path://M2 4.5 L22 4.5', '#336699', 24, 10, 2, 'diamond');

    expect(icon.tagName.toLowerCase()).toBe('svg');
    expect(icon.querySelector('path')?.getAttribute('stroke')).toBe('#336699');
    expect(icon.querySelectorAll('path')).toHaveLength(2);
  });

  it('does not render a marker when a legend item explicitly requests no icon', () => {
    const icon = createLegendIcon('none', '#336699', 12, 12);

    expect(icon === null).toBe(true);
  });

  it('builds anchored legend overlay from ECharts legend option', () => {
    const overlay = buildCustomLegendOverlay(
      {
        color: ['#ff0000'],
        legend: {
          bottom: '5%',
          orient: 'horizontal',
          data: ['Series A'],
          textStyle: { fontSize: 12, color: '#111111' },
        },
        series: [{ type: 'bar', data: [1] }],
      },
      { w: 400, h: 300 },
    );

    expect(overlay?.style.bottom).toBe('15px');
    expect(overlay?.textContent).toContain('Series A');
  });

  it('uses the series symbol size for line legend markers', () => {
    const overlay = buildCustomLegendOverlay(
      {
        color: ['#156082'],
        legend: {
          right: 20,
          top: 'middle',
          orient: 'vertical',
          itemWidth: 24,
          itemHeight: 9,
          data: [
            {
              name: 'Y value',
              icon: 'path://M2 4.5 L22 4.5',
              marker: 'diamond',
            },
          ],
          textStyle: { fontSize: 10, color: '#000000' },
        },
        series: [
          {
            type: 'line',
            name: 'Y value',
            symbolSize: 14,
            lineStyle: { color: '#156082', width: 2 },
            itemStyle: { color: '#156082' },
            data: [[0.7, 2.7]],
          },
        ],
      },
      { w: 800, h: 450 },
    );

    const svg = overlay?.querySelector('svg');
    const markerPath = svg?.querySelectorAll('path')[1];

    expect(svg?.getAttribute('height')).toBe('14');
    expect(markerPath?.getAttribute('d')).toBe('M12 0 L19 7 L12 14 L5 7 Z');
  });

  it('matches legend overlay colors by series name when legend order differs', () => {
    const overlay = buildCustomLegendOverlay(
      {
        color: ['#4472C4', '#ED7D31'],
        legend: {
          right: 20,
          top: 'middle',
          orient: 'vertical',
          itemWidth: 24,
          itemHeight: 9,
          data: [
            { name: 'Chat', icon: 'path://M2 4.5 L22 4.5' },
            { name: 'Email', icon: 'path://M2 4.5 L22 4.5' },
          ],
          textStyle: { fontSize: 10, color: '#000000' },
        },
        series: [
          {
            type: 'line',
            name: 'Email',
            lineStyle: { color: '#4472C4', width: 2 },
            itemStyle: { color: '#4472C4' },
            data: [120, 132],
          },
          {
            type: 'line',
            name: 'Chat',
            lineStyle: { color: '#ED7D31', width: 2 },
            itemStyle: { color: '#ED7D31' },
            data: [220, 182],
          },
        ],
      },
      { w: 800, h: 450 },
    );

    const icons = overlay?.querySelectorAll('svg path:first-child');
    expect(icons?.[0].getAttribute('stroke')).toBe('#ED7D31');
    expect(icons?.[1].getAttribute('stroke')).toBe('#4472C4');
  });

  it('shows the series picture fill in the legend swatch, not a palette color', () => {
    const overlay = buildCustomLegendOverlay(
      {
        color: ['#2f6f8f', '#ff8800'],
        legend: {
          bottom: '5%',
          orient: 'horizontal',
          data: ['系列 1'],
          textStyle: { fontSize: 14, color: '#AB0515' },
        },
        series: [
          {
            type: 'bar',
            name: '系列 1',
            data: [1],
            itemStyle: { color: { image: 'blob:pattern-url', repeat: 'repeat' } },
          },
        ],
      },
      { w: 400, h: 300 },
    );

    const image = overlay?.querySelector('image');
    expect(image?.getAttribute('href')).toBe('blob:pattern-url');
    expect(image?.getAttribute('preserveAspectRatio')).toBe('xMidYMid slice');
    // The palette color must not be painted underneath it.
    expect(overlay?.querySelector('rect')).toBeNull();
  });

  it('uses a gradient first stop for the swatch when the series is gradient filled', () => {
    const overlay = buildCustomLegendOverlay(
      {
        color: ['#2f6f8f'],
        legend: {
          bottom: '5%',
          orient: 'horizontal',
          data: ['Gradient'],
          textStyle: { fontSize: 10, color: '#000000' },
        },
        series: [
          {
            type: 'bar',
            name: 'Gradient',
            data: [1],
            itemStyle: {
              color: {
                colorStops: [
                  { offset: 0, color: 'rgba(171,5,21,1)' },
                  { offset: 1, color: 'rgba(255,255,240,1)' },
                ],
              },
            },
          },
        ],
      },
      { w: 400, h: 300 },
    );

    expect(overlay?.querySelector('rect')?.getAttribute('fill')).toBe('rgba(171,5,21,1)');
  });
});
