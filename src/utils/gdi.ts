/**
 * Shared GDI primitives for the metafile converters.
 *
 * EMF and WMF are two encodings of the same drawing model: the same pens,
 * brushes, stock objects, map modes and COLORREF layout, differing mainly in
 * record framing and integer width. Everything that describes *what* GDI draws
 * lives here so `emfVector` and `wmfVector` cannot drift apart on colour
 * conversion, map-mode scaling or stock-object defaults; each converter keeps
 * only its own record decoding.
 */

// --- Map modes --------------------------------------------------------------
const MM_TEXT = 1;
const MM_LOMETRIC = 2;
const MM_HIMETRIC = 3;
const MM_LOENGLISH = 4;
const MM_HIENGLISH = 5;
const MM_TWIPS = 6;
const MM_ISOTROPIC = 7;
const MM_ANISOTROPIC = 8;

/** Millimetres per logical unit for the fixed (non-scalable) map modes. */
const FIXED_MODE_MM_PER_UNIT: Record<number, number> = {
  [MM_LOMETRIC]: 0.1,
  [MM_HIMETRIC]: 0.01,
  [MM_LOENGLISH]: 0.254,
  [MM_HIENGLISH]: 0.0254,
  [MM_TWIPS]: 25.4 / 1440,
};

// --- Object handles ---------------------------------------------------------
export const STOCK_OBJECT_FLAG = 0x80000000;
const WHITE_BRUSH = 0x80000000;
const LTGRAY_BRUSH = 0x80000001;
const GRAY_BRUSH = 0x80000002;
const DKGRAY_BRUSH = 0x80000003;
const BLACK_BRUSH = 0x80000004;
const NULL_BRUSH = 0x80000005;
const WHITE_PEN = 0x80000006;
const BLACK_PEN = 0x80000007;
const NULL_PEN = 0x80000008;

export const BS_NULL = 1;
export const PS_STYLE_MASK = 0x0000000f;
export const PS_NULL = 5;

export const ALTERNATE_FILL = 1;

// --- Safety limits ----------------------------------------------------------
export const MAX_RECORDS = 500_000;
export const MAX_PATHS = 100_000;
export const MAX_POINTS_PER_RECORD = 100_000;

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A single filled and/or stroked contour, in device units. */
export interface MetafileVectorPath {
  /** SVG path data. */
  d: string;
  /** Fill color as `#RRGGBB`, or `none`. */
  fill: string;
  fillRule: 'nonzero' | 'evenodd';
  /** Stroke color as `#RRGGBB`, or `none`. */
  stroke: string;
  /** Stroke width in device units (0 means a hairline). */
  strokeWidth: number;
}

/** A converted metafile drawing. Coordinates are device units. */
export interface MetafileVectorImage {
  x: number;
  y: number;
  width: number;
  height: number;
  paths: MetafileVectorPath[];
}

export interface Point {
  x: number;
  y: number;
}

export interface Brush {
  /** False for BS_NULL (hollow) brushes. */
  visible: boolean;
  color: string;
}

export interface Pen {
  /** False for PS_NULL pens. */
  visible: boolean;
  color: string;
  /** Width in logical units; 0 means a cosmetic one-pixel pen. */
  width: number;
}

/** 2x3 affine matrix (GDI XFORM), mapping world space to page space. */
export interface Xform {
  m11: number;
  m12: number;
  m21: number;
  m22: number;
  dx: number;
  dy: number;
}

export const IDENTITY: Xform = { m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 };

export interface DeviceContext {
  brush: Brush;
  pen: Pen;
  fillRule: 'nonzero' | 'evenodd';
  mapMode: number;
  winOrg: Point;
  winExt: Point;
  vpOrg: Point;
  vpExt: Point;
  xform: Xform;
}

/** A pen and brush a freshly created DC starts with. */
export function createDeviceContext(): DeviceContext {
  return {
    brush: { visible: true, color: '#FFFFFF' },
    pen: { visible: true, color: '#000000', width: 0 },
    fillRule: 'evenodd',
    mapMode: MM_TEXT,
    winOrg: { x: 0, y: 0 },
    winExt: { x: 1, y: 1 },
    vpOrg: { x: 0, y: 0 },
    vpExt: { x: 1, y: 1 },
    xform: IDENTITY,
  };
}

/** Distinguishes the two object kinds stored in one GDI object table. */
export function isPen(obj: Brush | Pen): obj is Pen {
  return 'width' in obj;
}

// ---------------------------------------------------------------------------
// Coordinate mapping
// ---------------------------------------------------------------------------

/** Logical-to-device scale factors implied by the current map mode. */
export function mapScale(dc: DeviceContext, pxPerMm: Point | null): { sx: number; sy: number } {
  if (dc.mapMode === MM_ISOTROPIC || dc.mapMode === MM_ANISOTROPIC) {
    const sx = dc.winExt.x !== 0 ? dc.vpExt.x / dc.winExt.x : 1;
    const sy = dc.winExt.y !== 0 ? dc.vpExt.y / dc.winExt.y : 1;
    if (dc.mapMode === MM_ISOTROPIC) {
      // Isotropic keeps a 1:1 aspect ratio; GDI uses the smaller magnitude.
      const s = Math.min(Math.abs(sx), Math.abs(sy));
      return { sx: Math.sign(sx) * s || s, sy: Math.sign(sy) * s || s };
    }
    return { sx, sy };
  }

  const mmPerUnit = FIXED_MODE_MM_PER_UNIT[dc.mapMode];
  if (mmPerUnit !== undefined && pxPerMm) {
    // The metric modes have a y axis pointing up, unlike device space.
    return { sx: mmPerUnit * pxPerMm.x, sy: -mmPerUnit * pxPerMm.y };
  }

  return { sx: 1, sy: 1 }; // MM_TEXT and anything unrecognised
}

export function multiplyXform(a: Xform, b: Xform): Xform {
  return {
    m11: a.m11 * b.m11 + a.m12 * b.m21,
    m12: a.m11 * b.m12 + a.m12 * b.m22,
    m21: a.m21 * b.m11 + a.m22 * b.m21,
    m22: a.m21 * b.m12 + a.m22 * b.m22,
    dx: a.dx * b.m11 + a.dy * b.m21 + b.dx,
    dy: a.dx * b.m12 + a.dy * b.m22 + b.dy,
  };
}

// ---------------------------------------------------------------------------
// Geometry helpers
// ---------------------------------------------------------------------------

export function polyToSegments(
  points: Point[],
  closed: boolean,
  toDevice: (p: Point) => Point,
): string[] {
  const segments = [`M${fmtPoint(toDevice(points[0]))}`];
  for (let i = 1; i < points.length; i++) {
    segments.push(`L${fmtPoint(toDevice(points[i]))}`);
  }
  if (closed) segments.push('Z');
  return segments;
}

/** The four corners of a GDI bounding rectangle, clockwise from the top left. */
export function rectCorners(left: number, top: number, right: number, bottom: number): Point[] {
  return [
    { x: left, y: top },
    { x: right, y: top },
    { x: right, y: bottom },
    { x: left, y: bottom },
  ];
}

/** Approximate an axis-aligned ellipse inscribed in `corners` with four cubic arcs. */
export function ellipseSegments(corners: Point[], toDevice: (p: Point) => Point): string {
  const [tl, , br] = corners;
  const cx = (tl.x + br.x) / 2;
  const cy = (tl.y + br.y) / 2;
  const rx = (br.x - tl.x) / 2;
  const ry = (br.y - tl.y) / 2;
  const k = 0.5522847498307936; // 4/3 * (sqrt(2) - 1)
  const kx = rx * k;
  const ky = ry * k;
  const p = (x: number, y: number) => fmtPoint(toDevice({ x, y }));
  return (
    `M${p(cx - rx, cy)} ` +
    `C${p(cx - rx, cy - ky)} ${p(cx - kx, cy - ry)} ${p(cx, cy - ry)} ` +
    `C${p(cx + kx, cy - ry)} ${p(cx + rx, cy - ky)} ${p(cx + rx, cy)} ` +
    `C${p(cx + rx, cy + ky)} ${p(cx + kx, cy + ry)} ${p(cx, cy + ry)} ` +
    `C${p(cx - kx, cy + ry)} ${p(cx - rx, cy + ky)} ${p(cx - rx, cy)} Z`
  );
}

/**
 * A rounded rectangle, as GDI's RoundRect defines it: `rw`/`rh` are the *full*
 * width and height of the ellipse used to round each corner, so the corner radii
 * are half of them, clamped to the rectangle.
 */
export function roundRectSegments(
  corners: Point[],
  rw: number,
  rh: number,
  toDevice: (p: Point) => Point,
): string[] {
  const [tl, , br] = corners;
  const left = Math.min(tl.x, br.x);
  const right = Math.max(tl.x, br.x);
  const top = Math.min(tl.y, br.y);
  const bottom = Math.max(tl.y, br.y);
  const rx = Math.min(Math.abs(rw) / 2, (right - left) / 2);
  const ry = Math.min(Math.abs(rh) / 2, (bottom - top) / 2);
  if (rx <= 0 || ry <= 0) {
    return polyToSegments(rectCorners(left, top, right, bottom), true, toDevice);
  }

  const k = 0.5522847498307936;
  const kx = rx * k;
  const ky = ry * k;
  const p = (x: number, y: number) => fmtPoint(toDevice({ x, y }));
  return [
    `M${p(left + rx, top)}`,
    `L${p(right - rx, top)}`,
    `C${p(right - rx + kx, top)} ${p(right, top + ry - ky)} ${p(right, top + ry)}`,
    `L${p(right, bottom - ry)}`,
    `C${p(right, bottom - ry + ky)} ${p(right - rx + kx, bottom)} ${p(right - rx, bottom)}`,
    `L${p(left + rx, bottom)}`,
    `C${p(left + rx - kx, bottom)} ${p(left, bottom - ry + ky)} ${p(left, bottom - ry)}`,
    `L${p(left, top + ry)}`,
    `C${p(left, top + ry - ky)} ${p(left + rx - kx, top)} ${p(left + rx, top)}`,
    'Z',
  ];
}

/**
 * Sample the elliptical arc GDI's Arc/Pie/Chord records describe.
 *
 * GDI does not take sweep angles: it takes two radial rays from the ellipse
 * centre and draws counter-clockwise from the start ray to the end ray. The rays
 * only pick directions — their lengths are ignored. Sampling as a polyline keeps
 * the ellipse's own aspect ratio without converting visual angles to parametric
 * ones, which is what an SVG `A` command would require.
 */
export function arcPoints(corners: Point[], start: Point, end: Point, segments = 48): Point[] {
  const [tl, , br] = corners;
  const cx = (tl.x + br.x) / 2;
  const cy = (tl.y + br.y) / 2;
  const rx = Math.abs(br.x - tl.x) / 2;
  const ry = Math.abs(br.y - tl.y) / 2;
  if (rx === 0 || ry === 0) return [];

  // Map each ray onto the unit circle before taking its angle, so a point on an
  // elongated ellipse yields the parametric angle that actually lands on it.
  const angleOf = (p: Point) => Math.atan2((p.y - cy) / ry, (p.x - cx) / rx);
  const startAngle = angleOf(start);
  let sweep = angleOf(end) - startAngle;
  // GDI sweeps counter-clockwise, which is negative in a y-down device space.
  while (sweep > 0) sweep -= 2 * Math.PI;
  if (sweep === 0) sweep = -2 * Math.PI;

  const steps = Math.max(2, Math.ceil((Math.abs(sweep) / (2 * Math.PI)) * segments));
  const points: Point[] = [];
  for (let i = 0; i <= steps; i++) {
    const a = startAngle + (sweep * i) / steps;
    points.push({ x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) });
  }
  return points;
}

// ---------------------------------------------------------------------------
// Value formatting
// ---------------------------------------------------------------------------

/** A GDI COLORREF is `0x00bbggrr`. */
export function colorRefToHex(colorRef: number): string {
  const r = colorRef & 0xff;
  const g = (colorRef >> 8) & 0xff;
  const b = (colorRef >> 16) & 0xff;
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function toHex(value: number): string {
  return value.toString(16).padStart(2, '0').toUpperCase();
}

/** Round to 3 decimals and drop trailing zeroes, keeping the SVG compact. */
function fmtNumber(value: number): string {
  if (!Number.isFinite(value)) return '0';
  const rounded = Math.round(value * 1000) / 1000;
  return Object.is(rounded, -0) ? '0' : String(rounded);
}

export function fmtPoint(p: Point): string {
  return `${fmtNumber(p.x)},${fmtNumber(p.y)}`;
}

export function stockObject(handle: number): Brush | Pen | null {
  switch (handle >>> 0) {
    case WHITE_BRUSH:
      return { visible: true, color: '#FFFFFF' };
    case LTGRAY_BRUSH:
      return { visible: true, color: '#C0C0C0' };
    case GRAY_BRUSH:
      return { visible: true, color: '#808080' };
    case DKGRAY_BRUSH:
      return { visible: true, color: '#404040' };
    case BLACK_BRUSH:
      return { visible: true, color: '#000000' };
    case NULL_BRUSH:
      return { visible: false, color: '#000000' };
    case WHITE_PEN:
      return { visible: true, color: '#FFFFFF', width: 0 };
    case BLACK_PEN:
      return { visible: true, color: '#000000', width: 0 };
    case NULL_PEN:
      return { visible: false, color: '#000000', width: 0 };
    default:
      return null;
  }
}

// ---------------------------------------------------------------------------
// SVG serialization
// ---------------------------------------------------------------------------

/** Serialize a converted drawing as a standalone SVG document. */
export function metafileVectorToSvg(image: MetafileVectorImage): string {
  const body = image.paths
    .map((p) => {
      const attrs = [`d="${p.d}"`];
      attrs.push(`fill="${p.fill}"`);
      if (p.fill !== 'none' && p.fillRule === 'evenodd') attrs.push('fill-rule="evenodd"');
      if (p.stroke !== 'none') {
        attrs.push(`stroke="${p.stroke}"`);
        // A zero-width GDI pen is cosmetic: one device pixel.
        attrs.push(`stroke-width="${fmtNumber(p.strokeWidth || 1)}"`);
        attrs.push('stroke-linejoin="round"', 'stroke-linecap="round"');
      }
      return `<path ${attrs.join(' ')}/>`;
    })
    .join('');
  const viewBox =
    `${fmtNumber(image.x)} ${fmtNumber(image.y)} ` +
    `${fmtNumber(image.width)} ${fmtNumber(image.height)}`;
  // Explicit width/height give the document an intrinsic size, so it also works
  // as an <img> source or CSS background rather than defaulting to 300x150.
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmtNumber(image.width)}" ` +
    `height="${fmtNumber(image.height)}" viewBox="${viewBox}" ` +
    `preserveAspectRatio="none">${body}</svg>`
  );
}
