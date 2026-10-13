import { describe, expect, it } from 'vitest';
import {
  createFilledLineGeometry,
  createLineEndpointGeometry,
  insetLineSegment,
  roundedPolygonCommands,
} from '../../../src/shapes/editingGeometry';

const dimensions = { length: 24, width: 24, outlineWidth: 4 };

describe('framework-free editing geometry', () => {
  it('centers diamonds on their anchors and uses explicit dimensions', () => {
    expect(createLineEndpointGeometry('diamond', dimensions)).toEqual({
      kind: 'polygon',
      points: [
        [-12, 0],
        [0, -12],
        [12, 0],
        [0, 12],
      ],
    });
    expect(createLineEndpointGeometry('diamond', { ...dimensions, length: 40, width: 10 })).toEqual(
      {
        kind: 'polygon',
        points: [
          [-20, 0],
          [0, -5],
          [20, 0],
          [0, 5],
        ],
      },
    );
  });

  it('builds open arrowheads with a continuous positive-winding contour', () => {
    const geometry = createLineEndpointGeometry('arrow', dimensions);
    expect(geometry.kind).toBe('polygon');
    if (geometry.kind !== 'polygon') throw new Error('Expected polygon');
    expect(geometry.points).toHaveLength(6);
    const signedArea = geometry.points.reduce((area, [x, y], index, points) => {
      const [nextX, nextY] = points[(index + 1) % points.length];
      return area + x * nextY - nextX * y;
    }, 0);
    expect(signedArea).toBeGreaterThan(0);
    expect(geometry.points.every((point) => point.every(Number.isFinite))).toBe(true);
  });

  it('insets shafts without changing endpoint geometries or reversing short lines', () => {
    for (const [dx, dy] of [
      [300, 0],
      [0, 300],
      [-300, -100],
      [1, 0],
      [0, 0],
    ]) {
      const shaft = insetLineSegment(dx, dy, 22, 44);
      expect([...shaft.start, ...shaft.end].every(Number.isFinite)).toBe(true);
      expect(shaft.startInset + shaft.endInset).toBeLessThanOrEqual(Math.hypot(dx, dy) + 1e-8);
      expect(
        (shaft.end[0] - shaft.start[0]) * dx + (shaft.end[1] - shaft.start[1]) * dy,
      ).toBeGreaterThanOrEqual(0);
    }
  });

  it('keeps each filled endpoint path identical when the opposite endpoint changes', () => {
    const shapes = [
      'none',
      'circle',
      'square',
      'arrow',
      'triangle',
      'reversedTriangle',
      'diamond',
      'stealth',
    ] as const;
    for (const fixed of shapes) {
      const endpoint = createLineEndpointGeometry(fixed, dimensions);
      const none = createLineEndpointGeometry('none', dimensions);
      const baseline = createFilledLineGeometry({
        width: 300,
        height: 0,
        strokeWidth: 8,
        start: none,
        end: endpoint,
        startInset: 0,
        endInset: 8,
      });
      for (const opposite of shapes) {
        const geometry = createFilledLineGeometry({
          width: 300,
          height: 0,
          strokeWidth: 8,
          start: createLineEndpointGeometry(opposite, dimensions),
          end: endpoint,
          startInset: 4,
          endInset: 8,
        });
        expect([geometry.width, geometry.height]).toEqual([300, 1]);
        if (fixed !== 'none') expect(geometry.paths.at(-1)).toEqual(baseline.paths.at(-1));
      }
    }
  });

  it('uses caller-supplied polygons and clamps oversized corner radii', () => {
    for (const vertices of [
      [
        [0, 0],
        [100, 0],
        [100, 100],
        [0, 100],
      ],
      [
        [0, 0],
        [0, 0],
        [0, 100],
      ],
    ]) {
      const commands = roundedPolygonCommands(vertices as [number, number][], 1000);
      expect(commands.every((command) => command.points.every(Number.isFinite))).toBe(true);
      expect(commands.at(-1)?.kind).toBe('Z');
    }
    expect(roundedPolygonCommands([], 20)).toEqual([]);
  });

  it('rejects invalid dimensions and handles zero-size geometry without non-finite coordinates', () => {
    expect(() => createLineEndpointGeometry('triangle', { ...dimensions, length: NaN })).toThrow(
      RangeError,
    );
    expect(() => createLineEndpointGeometry('diamond', { ...dimensions, width: -1 })).toThrow(
      RangeError,
    );
    expect(() => insetLineSegment(100, 0, -1, 0)).toThrow(RangeError);
    const endpoint = createLineEndpointGeometry('arrow', { length: 0, width: 0, outlineWidth: 0 });
    const geometry = createFilledLineGeometry({
      width: 0,
      height: 0,
      strokeWidth: 0,
      start: endpoint,
      end: endpoint,
      startInset: 0,
      endInset: 0,
    });
    expect([geometry.width, geometry.height]).toEqual([1, 1]);
    expect(JSON.stringify(geometry)).not.toMatch(/null|NaN|Infinity/);
  });
});
