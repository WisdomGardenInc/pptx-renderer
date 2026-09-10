/**
 * EMF vector converter — turns the path/polygon subset of EMF (Enhanced Metafile)
 * drawing records into SVG geometry.
 *
 * Office writes plain vector EMF for decorative artwork (Chinese-style borders,
 * clip-art ornaments, converted shapes). Those files carry no embedded PDF and no
 * DIB, so `emfParser` classifies them as `unsupported` and nothing is drawn. In
 * practice they use a very narrow slice of GDI: solid brushes plus
 * BEGINPATH/…/FILLPATH built from MOVETOEX and POLYBEZIERTO16, which maps almost
 * one-to-one onto SVG paths.
 *
 * Supported: map modes, window/viewport mapping, world transforms, DC save/restore,
 * brush/pen objects (including stock objects), path bracket records, and the
 * polyline/polygon/bezier families in both 16-bit and 32-bit point form.
 *
 * Not supported (records are skipped, the rest of the drawing is still emitted):
 * text output, bitmap blits, clipping regions, arcs, gradients and hatch patterns.
 * The result is therefore a best-effort approximation, but a partial drawing beats
 * the blank frame the caller would otherwise show.
 */

// --- Record types -----------------------------------------------------------
const EMR_HEADER = 1;
const EMR_POLYBEZIER = 2;
const EMR_POLYGON = 3;
const EMR_POLYLINE = 4;
const EMR_POLYBEZIERTO = 5;
const EMR_POLYLINETO = 6;
const EMR_POLYPOLYLINE = 7;
const EMR_POLYPOLYGON = 8;
const EMR_SETWINDOWEXTEX = 9;
const EMR_SETWINDOWORGEX = 10;
const EMR_SETVIEWPORTEXTEX = 11;
const EMR_SETVIEWPORTORGEX = 12;
const EMR_EOF = 14;
const EMR_SETMAPMODE = 17;
const EMR_SETPOLYFILLMODE = 19;
const EMR_SAVEDC = 33;
const EMR_RESTOREDC = 34;
const EMR_SETWORLDTRANSFORM = 35;
const EMR_MODIFYWORLDTRANSFORM = 36;
const EMR_SELECTOBJECT = 37;
const EMR_CREATEPEN = 38;
const EMR_CREATEBRUSHINDIRECT = 39;
const EMR_DELETEOBJECT = 40;
const EMR_ELLIPSE = 42;
const EMR_RECTANGLE = 43;
const EMR_MOVETOEX = 27;
const EMR_LINETO = 54;
const EMR_BEGINPATH = 59;
const EMR_ENDPATH = 60;
const EMR_CLOSEFIGURE = 61;
const EMR_FILLPATH = 62;
const EMR_STROKEANDFILLPATH = 63;
const EMR_STROKEPATH = 64;
const EMR_ABORTPATH = 68;
const EMR_POLYBEZIER16 = 85;
const EMR_POLYGON16 = 86;
const EMR_POLYLINE16 = 87;
const EMR_POLYBEZIERTO16 = 88;
const EMR_POLYLINETO16 = 89;
const EMR_POLYPOLYLINE16 = 90;
const EMR_POLYPOLYGON16 = 91;
const EMR_EXTCREATEPEN = 95;

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
const STOCK_OBJECT_FLAG = 0x80000000;
const WHITE_BRUSH = 0x80000000;
const LTGRAY_BRUSH = 0x80000001;
const GRAY_BRUSH = 0x80000002;
const DKGRAY_BRUSH = 0x80000003;
const BLACK_BRUSH = 0x80000004;
const NULL_BRUSH = 0x80000005;
const WHITE_PEN = 0x80000006;
const BLACK_PEN = 0x80000007;
const NULL_PEN = 0x80000008;

const BS_NULL = 1;
const PS_STYLE_MASK = 0x0000000f;
const PS_NULL = 5;

const ALTERNATE_FILL = 1;

// --- Safety limits ----------------------------------------------------------
const MAX_RECORDS = 500_000;
const MAX_PATHS = 100_000;
const MAX_POINTS_PER_RECORD = 100_000;

/** A single filled and/or stroked contour, in EMF device units. */
export interface EmfVectorPath {
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

/** A converted EMF drawing. Coordinates are EMF device units. */
export interface EmfVectorImage {
  x: number;
  y: number;
  width: number;
  height: number;
  paths: EmfVectorPath[];
}

interface Point {
  x: number;
  y: number;
}

interface Brush {
  /** False for BS_NULL (hollow) brushes. */
  visible: boolean;
  color: string;
}

interface Pen {
  /** False for PS_NULL pens. */
  visible: boolean;
  color: string;
  /** Width in logical units; 0 means a cosmetic one-pixel pen. */
  width: number;
}

/** 2x3 affine matrix (GDI XFORM), mapping world space to page space. */
interface Xform {
  m11: number;
  m12: number;
  m21: number;
  m22: number;
  dx: number;
  dy: number;
}

const IDENTITY: Xform = { m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 };

interface DeviceContext {
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

/**
 * Parse an EMF file into SVG-ready vector geometry.
 *
 * Returns `null` when the data is not a valid EMF or when no drawable geometry
 * could be recovered, so callers can fall back to their existing behaviour.
 */
export function parseEmfVector(data: Uint8Array): EmfVectorImage | null {
  if (data.length < 88) return null;

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(40, true) !== 0x464d4520) return null; // " EMF"

  const headerSize = Math.min(view.getUint32(4, true), data.length);
  const frame = readViewBox(view, headerSize);
  if (!frame) return null;

  const paths: EmfVectorPath[] = [];
  let dc: DeviceContext = {
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
  const dcStack: DeviceContext[] = [];
  const objects = new Map<number, Brush | Pen>();

  // Device-space scale of the header, used to size fixed map modes.
  const pxPerMm = readPixelsPerMm(view, headerSize);

  /** Path currently being built between BEGINPATH and ENDPATH, as SVG segments. */
  let pathSegments: string[] = [];
  let inPathBracket = false;
  /** Current position in logical units. */
  let current: Point = { x: 0, y: 0 };
  let emitted = false;

  /** Map a logical point to device space through the world transform and map mode. */
  const toDevice = (p: Point): Point => {
    const { xform } = dc;
    const px = xform.m11 * p.x + xform.m21 * p.y + xform.dx;
    const py = xform.m12 * p.x + xform.m22 * p.y + xform.dy;
    const { sx, sy } = mapScale(dc, pxPerMm);
    return {
      x: (px - dc.winOrg.x) * sx + dc.vpOrg.x,
      y: (py - dc.winOrg.y) * sy + dc.vpOrg.y,
    };
  };

  /** Finish the pending geometry as one painted path. */
  const flush = (segments: string[], fill: boolean, stroke: boolean) => {
    if (segments.length === 0) return;
    if (paths.length >= MAX_PATHS) return;
    const filled = fill && dc.brush.visible;
    const stroked = stroke && dc.pen.visible;
    if (!filled && !stroked) return;
    const { sx, sy } = mapScale(dc, pxPerMm);
    const penScale = (Math.abs(sx) + Math.abs(sy)) / 2;
    paths.push({
      d: segments.join(' '),
      fill: filled ? dc.brush.color : 'none',
      fillRule: dc.fillRule,
      stroke: stroked ? dc.pen.color : 'none',
      strokeWidth: stroked ? Math.max(dc.pen.width * penScale, 0) : 0,
    });
    emitted = true;
  };

  /** Add geometry to the open path, or paint it straight away outside a bracket. */
  const draw = (segments: string[], fill: boolean, stroke: boolean) => {
    if (inPathBracket) {
      pathSegments.push(...segments);
      return;
    }
    flush(segments, fill, stroke);
  };

  /**
   * Add geometry from a `…To` record, which continues from the current position.
   * Outside a path bracket GDI strokes it immediately with the current pen.
   */
  const drawFromCurrent = (segments: string[], from: Point) => {
    if (inPathBracket) {
      if (pathSegments.length === 0) pathSegments.push(`M${fmtPoint(toDevice(from))}`);
      pathSegments.push(...segments);
      return;
    }
    flush([`M${fmtPoint(toDevice(from))}`, ...segments], false, true);
  };

  let offset = 0;
  let records = 0;

  while (offset + 8 <= data.length && records < MAX_RECORDS) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || size % 4 !== 0 || offset + size > data.length) break;
    records++;
    if (type === EMR_EOF) break;

    const body = offset + 8;
    const bodyBytes = size - 8;

    switch (type) {
      case EMR_HEADER:
        break;

      // --- Coordinate space ---
      case EMR_SETMAPMODE:
        if (bodyBytes >= 4) dc.mapMode = view.getUint32(body, true);
        break;
      case EMR_SETWINDOWORGEX:
        if (bodyBytes >= 8) dc.winOrg = readPointL(view, body);
        break;
      case EMR_SETWINDOWEXTEX:
        if (bodyBytes >= 8) dc.winExt = readPointL(view, body);
        break;
      case EMR_SETVIEWPORTORGEX:
        if (bodyBytes >= 8) dc.vpOrg = readPointL(view, body);
        break;
      case EMR_SETVIEWPORTEXTEX:
        if (bodyBytes >= 8) dc.vpExt = readPointL(view, body);
        break;
      case EMR_SETWORLDTRANSFORM:
        if (bodyBytes >= 24) dc.xform = readXform(view, body);
        break;
      case EMR_MODIFYWORLDTRANSFORM:
        if (bodyBytes >= 28) {
          const xf = readXform(view, body);
          const mode = view.getUint32(body + 24, true);
          // MWT_IDENTITY=1, MWT_LEFTMULTIPLY=2, MWT_RIGHTMULTIPLY=3
          if (mode === 1) dc.xform = IDENTITY;
          else if (mode === 2) dc.xform = multiplyXform(xf, dc.xform);
          else if (mode === 3) dc.xform = multiplyXform(dc.xform, xf);
        }
        break;

      // --- Device context stack ---
      case EMR_SAVEDC:
        dcStack.push({ ...dc });
        break;
      case EMR_RESTOREDC: {
        // SavedDC is a relative level: -1 pops one saved state, -2 pops two, …
        const level = bodyBytes >= 4 ? view.getInt32(body, true) : -1;
        const pops = level < 0 ? Math.min(-level, dcStack.length) : dcStack.length;
        for (let i = 0; i < pops; i++) {
          const restored = dcStack.pop();
          if (restored) dc = restored;
        }
        break;
      }

      // --- Fill rule ---
      case EMR_SETPOLYFILLMODE:
        if (bodyBytes >= 4) {
          dc.fillRule = view.getUint32(body, true) === ALTERNATE_FILL ? 'evenodd' : 'nonzero';
        }
        break;

      // --- Objects ---
      case EMR_CREATEBRUSHINDIRECT:
        if (bodyBytes >= 16) {
          const handle = view.getUint32(body, true);
          const style = view.getUint32(body + 4, true);
          objects.set(handle, {
            visible: style !== BS_NULL,
            color: colorRefToHex(view.getUint32(body + 8, true)),
          });
        }
        break;
      case EMR_CREATEPEN:
        if (bodyBytes >= 20) {
          const handle = view.getUint32(body, true);
          const style = view.getUint32(body + 4, true);
          objects.set(handle, {
            visible: (style & PS_STYLE_MASK) !== PS_NULL,
            color: colorRefToHex(view.getUint32(body + 16, true)),
            width: Math.abs(view.getInt32(body + 8, true)),
          });
        }
        break;
      case EMR_EXTCREATEPEN:
        // ihPen(4) offBmi(4) cbBmi(4) offBits(4) cbBits(4) then ELP: style(4) width(4)
        // brushStyle(4) color(4) …
        if (bodyBytes >= 36) {
          const handle = view.getUint32(body, true);
          const style = view.getUint32(body + 20, true);
          objects.set(handle, {
            visible: (style & PS_STYLE_MASK) !== PS_NULL,
            color: colorRefToHex(view.getUint32(body + 32, true)),
            width: Math.abs(view.getInt32(body + 24, true)),
          });
        }
        break;
      case EMR_SELECTOBJECT:
        if (bodyBytes >= 4) {
          const handle = view.getUint32(body, true);
          if (handle & STOCK_OBJECT_FLAG) {
            const stock = stockObject(handle);
            if (stock) {
              if ('width' in stock) dc.pen = stock;
              else dc.brush = stock;
            }
          } else {
            const obj = objects.get(handle);
            if (obj) {
              if ('width' in obj) dc.pen = obj;
              else dc.brush = obj;
            }
          }
        }
        break;
      case EMR_DELETEOBJECT:
        if (bodyBytes >= 4) objects.delete(view.getUint32(body, true));
        break;

      // --- Path bracket ---
      case EMR_BEGINPATH:
        inPathBracket = true;
        pathSegments = [];
        break;
      case EMR_ENDPATH:
        inPathBracket = false;
        break;
      case EMR_ABORTPATH:
        inPathBracket = false;
        pathSegments = [];
        break;
      case EMR_CLOSEFIGURE:
        if (pathSegments.length > 0) pathSegments.push('Z');
        break;
      case EMR_FILLPATH:
      case EMR_STROKEANDFILLPATH:
      case EMR_STROKEPATH:
        flush(pathSegments, type !== EMR_STROKEPATH, type !== EMR_FILLPATH);
        pathSegments = [];
        break;

      // --- Geometry ---
      case EMR_MOVETOEX:
        if (bodyBytes >= 8) {
          current = readPointL(view, body);
          // Outside a path bracket MoveTo only relocates the current position.
          if (inPathBracket) pathSegments.push(`M${fmtPoint(toDevice(current))}`);
        }
        break;
      case EMR_LINETO:
        if (bodyBytes >= 8) {
          const to = readPointL(view, body);
          drawFromCurrent([`L${fmtPoint(toDevice(to))}`], current);
          current = to;
        }
        break;

      case EMR_POLYBEZIERTO:
      case EMR_POLYBEZIERTO16: {
        const pts = readPolyPoints(view, body, bodyBytes, type === EMR_POLYBEZIERTO16);
        if (!pts || pts.length < 3) break;
        const segments: string[] = [];
        for (let i = 0; i + 2 < pts.length; i += 3) {
          segments.push(
            `C${fmtPoint(toDevice(pts[i]))} ${fmtPoint(toDevice(pts[i + 1]))} ` +
              `${fmtPoint(toDevice(pts[i + 2]))}`,
          );
        }
        drawFromCurrent(segments, current);
        current = pts[pts.length - 1];
        break;
      }
      case EMR_POLYLINETO:
      case EMR_POLYLINETO16: {
        const pts = readPolyPoints(view, body, bodyBytes, type === EMR_POLYLINETO16);
        if (!pts) break;
        drawFromCurrent(
          pts.map((p) => `L${fmtPoint(toDevice(p))}`),
          current,
        );
        current = pts[pts.length - 1];
        break;
      }
      case EMR_POLYBEZIER:
      case EMR_POLYBEZIER16: {
        const pts = readPolyPoints(view, body, bodyBytes, type === EMR_POLYBEZIER16);
        if (!pts || pts.length === 0) break;
        const segments: string[] = [];
        segments.push(`M${fmtPoint(toDevice(pts[0]))}`);
        for (let i = 1; i + 2 < pts.length; i += 3) {
          segments.push(
            `C${fmtPoint(toDevice(pts[i]))} ${fmtPoint(toDevice(pts[i + 1]))} ` +
              `${fmtPoint(toDevice(pts[i + 2]))}`,
          );
        }
        current = pts[pts.length - 1];
        draw(segments, false, true);
        break;
      }
      case EMR_POLYGON:
      case EMR_POLYGON16:
      case EMR_POLYLINE:
      case EMR_POLYLINE16: {
        const is16 = type === EMR_POLYGON16 || type === EMR_POLYLINE16;
        const closed = type === EMR_POLYGON || type === EMR_POLYGON16;
        const pts = readPolyPoints(view, body, bodyBytes, is16);
        if (!pts || pts.length === 0) break;
        const segments = polyToSegments(pts, closed, toDevice);
        current = pts[pts.length - 1];
        draw(segments, closed, true);
        break;
      }
      case EMR_POLYPOLYGON:
      case EMR_POLYPOLYGON16:
      case EMR_POLYPOLYLINE:
      case EMR_POLYPOLYLINE16: {
        const is16 = type === EMR_POLYPOLYGON16 || type === EMR_POLYPOLYLINE16;
        const closed = type === EMR_POLYPOLYGON || type === EMR_POLYPOLYGON16;
        const polygons = readPolyPolyPoints(view, body, bodyBytes, is16);
        if (!polygons) break;
        const segments: string[] = [];
        for (const poly of polygons) {
          if (poly.length === 0) continue;
          segments.push(...polyToSegments(poly, closed, toDevice));
          current = poly[poly.length - 1];
        }
        draw(segments, closed, true);
        break;
      }

      case EMR_RECTANGLE:
      case EMR_ELLIPSE: {
        if (bodyBytes < 16) break;
        const left = view.getInt32(body, true);
        const top = view.getInt32(body + 4, true);
        const right = view.getInt32(body + 8, true);
        const bottom = view.getInt32(body + 12, true);
        const corners = [
          { x: left, y: top },
          { x: right, y: top },
          { x: right, y: bottom },
          { x: left, y: bottom },
        ];
        if (type === EMR_RECTANGLE) {
          draw(polyToSegments(corners, true, toDevice), true, true);
        } else {
          draw([ellipseSegments(corners, toDevice)], true, true);
        }
        break;
      }

      default:
        // Unsupported record (text, blit, clip, arc, …) — skip it and keep going.
        break;
    }

    offset += size;
  }

  if (!emitted) return null;
  return { ...frame, paths };
}

/** Serialize a converted drawing as a standalone SVG document. */
export function emfVectorToSvg(image: EmfVectorImage): string {
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

// ---------------------------------------------------------------------------
// Header helpers
// ---------------------------------------------------------------------------

/**
 * Derive the drawing's viewBox in device units.
 *
 * `rclFrame` (in 0.01 mm) is the rectangle Office maps onto the destination, so it
 * is preferred; `rclBounds` (already in device units) is the fallback when the
 * frame or the device metrics are degenerate.
 */
function readViewBox(view: DataView, headerSize: number): Omit<EmfVectorImage, 'paths'> | null {
  const bounds = {
    left: view.getInt32(8, true),
    top: view.getInt32(12, true),
    right: view.getInt32(16, true),
    bottom: view.getInt32(20, true),
  };

  {
    const frame = {
      left: view.getInt32(24, true),
      top: view.getInt32(28, true),
      right: view.getInt32(32, true),
      bottom: view.getInt32(36, true),
    };
    const pxPerMm = readPixelsPerMm(view, headerSize);
    if (pxPerMm && frame.right > frame.left && frame.bottom > frame.top) {
      const x = (frame.left / 100) * pxPerMm.x;
      const y = (frame.top / 100) * pxPerMm.y;
      return {
        x,
        y,
        width: (frame.right / 100) * pxPerMm.x - x,
        height: (frame.bottom / 100) * pxPerMm.y - y,
      };
    }
  }

  if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) return null;
  return {
    x: bounds.left,
    y: bounds.top,
    width: bounds.right - bounds.left,
    height: bounds.bottom - bounds.top,
  };
}

/** Device pixels per millimetre, from the header's device/millimetre sizes. */
function readPixelsPerMm(view: DataView, headerSize: number): Point | null {
  // szlDevice/szlMillimeters live at offsets 72..88 of the header record.
  if (headerSize < 88 || view.byteLength < 88) return null;
  const devX = view.getInt32(72, true);
  const devY = view.getInt32(76, true);
  const mmX = view.getInt32(80, true);
  const mmY = view.getInt32(84, true);
  if (devX <= 0 || devY <= 0 || mmX <= 0 || mmY <= 0) return null;
  return { x: devX / mmX, y: devY / mmY };
}

// ---------------------------------------------------------------------------
// Coordinate mapping
// ---------------------------------------------------------------------------

/** Logical-to-device scale factors implied by the current map mode. */
function mapScale(dc: DeviceContext, pxPerMm: Point | null): { sx: number; sy: number } {
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

function multiplyXform(a: Xform, b: Xform): Xform {
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
// Record readers
// ---------------------------------------------------------------------------

function readPointL(view: DataView, offset: number): Point {
  return { x: view.getInt32(offset, true), y: view.getInt32(offset + 4, true) };
}

function readXform(view: DataView, offset: number): Xform {
  return {
    m11: view.getFloat32(offset, true),
    m12: view.getFloat32(offset + 4, true),
    m21: view.getFloat32(offset + 8, true),
    m22: view.getFloat32(offset + 12, true),
    dx: view.getFloat32(offset + 16, true),
    dy: view.getFloat32(offset + 20, true),
  };
}

/**
 * Read a `Bounds(16) Count(4) aPoints[]` payload shared by the poly* records.
 */
function readPolyPoints(
  view: DataView,
  body: number,
  bodyBytes: number,
  is16Bit: boolean,
): Point[] | null {
  if (bodyBytes < 20) return null;
  const count = view.getUint32(body + 16, true);
  if (count === 0 || count > MAX_POINTS_PER_RECORD) return null;
  const stride = is16Bit ? 4 : 8;
  if (20 + count * stride > bodyBytes) return null;
  return readPointArray(view, body + 20, count, is16Bit);
}

/**
 * Read a `Bounds(16) NumberOfPolys(4) Count(4) PolyPointCount[] aPoints[]` payload.
 */
function readPolyPolyPoints(
  view: DataView,
  body: number,
  bodyBytes: number,
  is16Bit: boolean,
): Point[][] | null {
  if (bodyBytes < 24) return null;
  const polyCount = view.getUint32(body + 16, true);
  const totalPoints = view.getUint32(body + 20, true);
  if (polyCount === 0 || polyCount > 65536) return null;
  if (totalPoints === 0 || totalPoints > MAX_POINTS_PER_RECORD) return null;

  const countsOffset = body + 24;
  const stride = is16Bit ? 4 : 8;
  if (24 + polyCount * 4 + totalPoints * stride > bodyBytes) return null;

  const pointsOffset = countsOffset + polyCount * 4;
  const polygons: Point[][] = [];
  let consumed = 0;
  for (let i = 0; i < polyCount; i++) {
    const n = view.getUint32(countsOffset + i * 4, true);
    if (n === 0 || consumed + n > totalPoints) break;
    polygons.push(readPointArray(view, pointsOffset + consumed * stride, n, is16Bit));
    consumed += n;
  }
  return polygons.length > 0 ? polygons : null;
}

function readPointArray(view: DataView, offset: number, count: number, is16Bit: boolean): Point[] {
  const points: Point[] = new Array(count);
  for (let i = 0; i < count; i++) {
    points[i] = is16Bit
      ? { x: view.getInt16(offset + i * 4, true), y: view.getInt16(offset + i * 4 + 2, true) }
      : { x: view.getInt32(offset + i * 8, true), y: view.getInt32(offset + i * 8 + 4, true) };
  }
  return points;
}

// ---------------------------------------------------------------------------
// Geometry helpers
// ---------------------------------------------------------------------------

function polyToSegments(points: Point[], closed: boolean, toDevice: (p: Point) => Point): string[] {
  const segments = [`M${fmtPoint(toDevice(points[0]))}`];
  for (let i = 1; i < points.length; i++) {
    segments.push(`L${fmtPoint(toDevice(points[i]))}`);
  }
  if (closed) segments.push('Z');
  return segments;
}

/** Approximate an axis-aligned ellipse inscribed in `corners` with four cubic arcs. */
function ellipseSegments(corners: Point[], toDevice: (p: Point) => Point): string {
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

// ---------------------------------------------------------------------------
// Value formatting
// ---------------------------------------------------------------------------

/** A GDI COLORREF is `0x00bbggrr`. */
function colorRefToHex(colorRef: number): string {
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

function fmtPoint(p: Point): string {
  return `${fmtNumber(p.x)},${fmtNumber(p.y)}`;
}

function stockObject(handle: number): Brush | Pen | null {
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
