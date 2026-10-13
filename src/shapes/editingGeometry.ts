export type ShapePoint = readonly [number, number];
export type ShapePathCommand = {
  readonly kind: 'M' | 'L' | 'Q' | 'C' | 'Z';
  readonly points: readonly number[];
};
export interface ShapeArcCommand {
  readonly kind: 'arc';
  readonly radiusX: number;
  readonly radiusY: number;
  readonly startAngle: number;
  readonly sweepAngle: number;
}
export type ShapeGeometryCommand = ShapePathCommand | ShapeArcCommand;
export interface CustomShapePath {
  readonly commands: readonly ShapeGeometryCommand[];
  readonly fill?: 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  readonly stroke?: boolean;
}
export interface CustomShapeGeometry {
  readonly width: number;
  readonly height: number;
  readonly paths: readonly CustomShapePath[];
}
export type LineEndpointShape =
  | 'none'
  | 'circle'
  | 'square'
  | 'arrow'
  | 'triangle'
  | 'reversedTriangle'
  | 'diamond'
  | 'stealth';
export type LineEndpointGeometry =
  | { readonly kind: 'none' }
  | { readonly kind: 'circle'; readonly radius: number }
  | { readonly kind: 'polygon'; readonly points: readonly ShapePoint[] };
export interface LineEndpointDimensions {
  readonly length: number;
  readonly width: number;
  readonly outlineWidth: number;
}
export interface LineSegmentGeometry {
  readonly start: ShapePoint;
  readonly end: ShapePoint;
  readonly startInset: number;
  readonly endInset: number;
}
export interface FilledLineOptions {
  readonly width: number;
  readonly height: number;
  readonly strokeWidth: number;
  readonly start: LineEndpointGeometry;
  readonly end: LineEndpointGeometry;
  readonly startInset: number;
  readonly endInset: number;
}

function nonnegative(...values: number[]): void {
  if (values.some((value) => !Number.isFinite(value) || value < 0))
    throw new RangeError('Geometry dimensions must be finite and nonnegative.');
}

export function polygonCommands(points: readonly ShapePoint[]): ShapePathCommand[] {
  if (!points.length) return [];
  return [
    ...points.map(
      (point, index): ShapePathCommand => ({ kind: index === 0 ? 'M' : 'L', points: point }),
    ),
    { kind: 'Z', points: [] },
  ];
}

export function roundedPolygonCommands(
  vertices: readonly ShapePoint[],
  radius: number,
): ShapePathCommand[] {
  nonnegative(radius);
  if (vertices.some((point) => !point.every(Number.isFinite)))
    throw new RangeError('Polygon coordinates must be finite.');
  if (!vertices.length) return [];
  const corners = vertices.map((point, index) => {
    const previous = vertices[(index + vertices.length - 1) % vertices.length];
    const next = vertices[(index + 1) % vertices.length];
    const before = Math.hypot(previous[0] - point[0], previous[1] - point[1]);
    const after = Math.hypot(next[0] - point[0], next[1] - point[1]);
    const cosine =
      ((previous[0] - point[0]) * (next[0] - point[0]) +
        (previous[1] - point[1]) * (next[1] - point[1])) /
      (before * after);
    const tangent = Math.tan(Math.acos(Math.max(-1, Math.min(1, cosine))) / 2);
    const distance = Math.min(radius / tangent, before / 2, after / 2);
    if (!Number.isFinite(distance) || before === 0 || after === 0)
      return { point, enter: point, exit: point };
    const toward = (target: ShapePoint, length: number): ShapePoint => [
      point[0] + ((target[0] - point[0]) * distance) / length,
      point[1] + ((target[1] - point[1]) * distance) / length,
    ];
    return { point, enter: toward(previous, before), exit: toward(next, after) };
  });
  const commands: ShapePathCommand[] = [{ kind: 'M', points: corners[0].enter }];
  for (const corner of corners)
    commands.push(
      { kind: 'L', points: corner.enter },
      { kind: 'Q', points: [...corner.point, ...corner.exit] },
    );
  commands.push({ kind: 'Z', points: [] });
  return commands;
}

export function shapeCommandsToSvgPath(commands: readonly ShapePathCommand[]): string {
  return commands.map((command) => `${command.kind}${command.points.join(' ')}`).join(' ');
}

export function openArrowheadPoints(
  length: number,
  width: number,
  outlineWidth: number,
): ShapePoint[] {
  nonnegative(length, width, outlineWidth);
  const half = width / 2;
  const diagonal = Math.hypot(length, half);
  const thickness = outlineWidth / 2;
  const ox = diagonal > 0 ? (thickness * half) / diagonal : 0;
  const oy = diagonal > 0 ? (thickness * length) / diagonal : 0;
  const miter = half > 0 ? (thickness * diagonal) / half : 0;
  return [
    [length - ox, -half - oy],
    [length + ox, -half + oy],
    [miter, 0],
    [length + ox, half - oy],
    [length - ox, half + oy],
    [-miter, 0],
  ];
}

export function createLineEndpointGeometry(
  shape: LineEndpointShape,
  dimensions: LineEndpointDimensions,
): LineEndpointGeometry {
  const { length, width, outlineWidth } = dimensions;
  nonnegative(length, width, outlineWidth);
  if (shape === 'none') return { kind: 'none' };
  if (shape === 'circle') return { kind: 'circle', radius: width / 2 };
  if (shape === 'arrow')
    return { kind: 'polygon', points: openArrowheadPoints(length, width, outlineWidth) };
  const x = length / 2;
  const y = width / 2;
  const polygons: Record<
    Exclude<LineEndpointShape, 'none' | 'circle' | 'arrow'>,
    readonly ShapePoint[]
  > = {
    square: [
      [-x, -y],
      [x, -y],
      [x, y],
      [-x, y],
    ],
    triangle: [
      [0, 0],
      [length, -y],
      [length, y],
    ],
    reversedTriangle: [
      [0, -y],
      [length, 0],
      [0, y],
    ],
    diamond: [
      [-x, 0],
      [0, -y],
      [x, 0],
      [0, y],
    ],
    stealth: [
      [0, 0],
      [length, -y],
      [(length * 2) / 3, 0],
      [length, y],
    ],
  };
  return { kind: 'polygon', points: polygons[shape] };
}

export function insetLineSegment(
  dx: number,
  dy: number,
  startOffset: number,
  endOffset: number,
): LineSegmentGeometry {
  if (![dx, dy].every(Number.isFinite)) throw new RangeError('Line coordinates must be finite.');
  nonnegative(startOffset, endOffset);
  const length = Math.hypot(dx, dy);
  const total = startOffset + endOffset;
  const ratio = total > 0 ? Math.min(1, length / total) : 1;
  const startInset = startOffset * ratio;
  const endInset = endOffset * ratio;
  const ux = length > 0 ? dx / length : 0;
  const uy = length > 0 ? dy / length : 0;
  const start: ShapePoint = [ux * startInset, uy * startInset];
  const end: ShapePoint = total >= length ? start : [dx - ux * endInset, dy - uy * endInset];
  return { start, end, startInset, endInset };
}

function filledSegmentCommands(
  start: ShapePoint,
  end: ShapePoint,
  thickness: number,
): ShapePathCommand[] {
  const dx = end[0] - start[0];
  const dy = end[1] - start[1];
  const length = Math.hypot(dx, dy);
  const nx = length > 0 ? ((-dy / length) * thickness) / 2 : 0;
  const ny = length > 0 ? ((dx / length) * thickness) / 2 : thickness / 2;
  return polygonCommands([
    [start[0] - nx, start[1] - ny],
    [end[0] - nx, end[1] - ny],
    [end[0] + nx, end[1] + ny],
    [start[0] + nx, start[1] + ny],
  ]);
}

function endpointCommands(
  geometry: LineEndpointGeometry,
  x: number,
  y: number,
  dx: number,
  dy: number,
): ShapeGeometryCommand[] {
  if (geometry.kind === 'none') return [];
  if (geometry.kind === 'circle')
    return [
      { kind: 'M', points: [x + geometry.radius, y] },
      {
        kind: 'arc',
        radiusX: geometry.radius,
        radiusY: geometry.radius,
        startAngle: 0,
        sweepAngle: 180,
      },
      {
        kind: 'arc',
        radiusX: geometry.radius,
        radiusY: geometry.radius,
        startAngle: 180,
        sweepAngle: 180,
      },
      { kind: 'Z', points: [] },
    ];
  const length = Math.hypot(dx, dy);
  const ux = length > 0 ? dx / length : 1;
  const uy = length > 0 ? dy / length : 0;
  return polygonCommands(
    geometry.points.map(([along, across]) => [
      x + along * ux - across * uy,
      y + along * uy + across * ux,
    ]),
  );
}

export function createFilledLineGeometry(options: FilledLineOptions): CustomShapeGeometry {
  const { width, height, strokeWidth, start, end, startInset, endInset } = options;
  nonnegative(width, height, strokeWidth);
  const shaft = insetLineSegment(width, height, startInset, endInset);
  const commands = [
    filledSegmentCommands(shaft.start, shaft.end, strokeWidth),
    endpointCommands(start, 0, 0, width, height),
    endpointCommands(end, width, height, -width, -height),
  ];
  return {
    width: Math.max(1, width),
    height: Math.max(1, height),
    paths: commands
      .filter((path) => path.length > 0)
      .map((path) => ({ commands: path, stroke: false, fill: 'norm' })),
  };
}
