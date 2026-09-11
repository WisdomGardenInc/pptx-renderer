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
 *
 * GDI state, colour conversion and SVG serialization are shared with the WMF
 * converter in `gdi.ts`; this module only decodes EMF's own record framing.
 */

import {
  ALTERNATE_FILL,
  BS_NULL,
  Brush,
  DeviceContext,
  IDENTITY,
  MAX_PATHS,
  MAX_POINTS_PER_RECORD,
  MAX_RECORDS,
  MetafileVectorImage,
  PS_NULL,
  PS_STYLE_MASK,
  Pen,
  Point,
  STOCK_OBJECT_FLAG,
  Xform,
  colorRefToHex,
  createDeviceContext,
  ellipseSegments,
  fmtPoint,
  isPen,
  mapScale,
  metafileVectorToSvg,
  multiplyXform,
  polyToSegments,
  rectCorners,
  stockObject,
} from './gdi';

/** A converted EMF drawing. Coordinates are EMF device units. */
export type EmfVectorImage = MetafileVectorImage;

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

export function parseEmfVector(data: Uint8Array): EmfVectorImage | null {
  if (data.length < 88) return null;

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(40, true) !== 0x464d4520) return null; // " EMF"

  const headerSize = Math.min(view.getUint32(4, true), data.length);
  const frame = readViewBox(view, headerSize);
  if (!frame) return null;

  const paths: EmfVectorImage['paths'] = [];
  let dc: DeviceContext = createDeviceContext();
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
              if (isPen(stock)) dc.pen = stock;
              else dc.brush = stock;
            }
          } else {
            const obj = objects.get(handle);
            if (obj) {
              if (isPen(obj)) dc.pen = obj;
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
        const corners = rectCorners(left, top, right, bottom);
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
export const emfVectorToSvg = metafileVectorToSvg;

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
