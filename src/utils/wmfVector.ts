/**
 * WMF vector converter — turns Windows Metafile drawing records into SVG geometry.
 *
 * WMF is the 16-bit ancestor of EMF and shows up in PPTX decks that carry older
 * clip art, artwork pasted in from Word, or shapes exported by pre-2007 Office.
 * Browsers cannot display it at all, so without this the picture is a hole in the
 * slide.
 *
 * It draws with the same GDI model as EMF — the pens, brushes, stock objects,
 * map modes and COLORREF layout in `gdi.ts` are shared — but the encoding differs
 * in four ways that drive everything below:
 *
 *  1. Record lengths are counted in 16-bit WORDs, not bytes, and the function
 *     code sits *after* the length rather than before it.
 *  2. Coordinates are INT16, and the ones in rectangle-shaped records are stored
 *     **in reverse**: `bottom, right, top, left`. Points inside the poly* arrays
 *     are the exception and stay in natural `x, y` order.
 *  3. There is no path bracket and no bezier record. Every figure is drawn by a
 *     single self-contained record, so geometry is emitted as it is read.
 *  4. Objects have no handles. Each create record takes the lowest free slot in
 *     an object table and later records address it by index, so create records
 *     this converter does not otherwise care about (fonts, palettes, regions)
 *     still have to consume their slot or every later index is off by one.
 *
 * Supported: placeable and plain headers, map modes, window/viewport mapping,
 * DC save/restore, pens and brushes including stock objects, polygon/polyline/
 * polypolygon, rectangle, round rectangle, ellipse, arc/pie/chord, line/move,
 * and text output with fonts, charsets, colour and alignment.
 *
 * Not supported (records are skipped, the rest of the drawing is still emitted):
 * bitmap blits, clipping regions, hatch and pattern brush textures (drawn as
 * their solid colour), opaque text backgrounds, and raster operation modes.
 */

import {
  ALTERNATE_FILL,
  BS_NULL,
  Brush,
  DeviceContext,
  GdiFont,
  MAX_PATHS,
  MAX_POINTS_PER_RECORD,
  MAX_RECORDS,
  MAX_TEXTS,
  MetafileVectorImage,
  PS_NULL,
  PS_STYLE_MASK,
  Pen,
  Point,
  STOCK_OBJECT_FLAG,
  TA_UPDATECP,
  arcPoints,
  baselineOffset,
  colorRefToHex,
  createDeviceContext,
  decodeGdiFacename,
  decodeGdiText,
  ellipseSegments,
  estimateTextAdvance,
  fmtPoint,
  gdiFontStack,
  isFont,
  isPen,
  logFontEmSize,
  mapScale,
  metafileVectorToSvg,
  polyToSegments,
  rectCorners,
  roundRectSegments,
  stockObject,
  textAnchorOf,
} from './gdi';

// --- Headers ----------------------------------------------------------------
/** Magic number introducing the 22-byte Aldus placeable header. */
const PLACEABLE_KEY = 0x9ac6cdd7;
const PLACEABLE_HEADER_SIZE = 22;
/** METAHEADER is 18 bytes, i.e. 9 WORDs — the value its HeaderSize field must hold. */
const META_HEADER_WORDS = 9;
const META_HEADER_SIZE = 18;

// --- Record function codes --------------------------------------------------
const META_EOF = 0x0000;
const META_SETBKMODE = 0x0102;
const META_SETMAPMODE = 0x0103;
const META_SETROP2 = 0x0104;
const META_SETPOLYFILLMODE = 0x0106;
const META_SETSTRETCHBLTMODE = 0x0107;
const META_SETTEXTCHAREXTRA = 0x0108;
const META_SETTEXTALIGN = 0x012e;
const META_SETTEXTCOLOR = 0x0209;
const META_TEXTOUT = 0x0521;
const META_EXTTEXTOUT = 0x0a32;
const META_SAVEDC = 0x001e;
const META_RESTOREDC = 0x0127;
const META_SELECTOBJECT = 0x012d;
const META_DELETEOBJECT = 0x01f0;
const META_SETWINDOWORG = 0x020b;
const META_SETWINDOWEXT = 0x020c;
const META_SETVIEWPORTORG = 0x020d;
const META_SETVIEWPORTEXT = 0x020e;
const META_OFFSETWINDOWORG = 0x020f;
const META_OFFSETVIEWPORTORG = 0x0211;
const META_LINETO = 0x0213;
const META_MOVETO = 0x0214;
const META_SCALEWINDOWEXT = 0x0410;
const META_SCALEVIEWPORTEXT = 0x0412;
const META_POLYGON = 0x0324;
const META_POLYLINE = 0x0325;
const META_POLYPOLYGON = 0x0538;
const META_RECTANGLE = 0x041b;
const META_ROUNDRECT = 0x061c;
const META_ELLIPSE = 0x0418;
const META_ARC = 0x0817;
const META_PIE = 0x081a;
const META_CHORD = 0x0830;
const META_CREATEPENINDIRECT = 0x02fa;
const META_CREATEBRUSHINDIRECT = 0x02fc;
const META_CREATEFONTINDIRECT = 0x02fb;
const META_CREATEPALETTE = 0x00f7;
const META_CREATEPATTERNBRUSH = 0x01f9;
const META_CREATEREGION = 0x06ff;
const META_DIBCREATEPATTERNBRUSH = 0x0142;

/**
 * Create records this converter cannot turn into a pen or brush, but which still
 * claim an object-table slot. Tracking them keeps later SELECTOBJECT indices right.
 */
const OPAQUE_CREATE_RECORDS = new Set([
  META_CREATEPALETTE,
  META_CREATEPATTERNBRUSH,
  META_CREATEREGION,
  META_DIBCREATEPATTERNBRUSH,
]);

/**
 * A slot held by an object this converter cannot draw with — a font, palette or
 * region. It has to stay distinct from a free slot: GDI hands each create record
 * the lowest *free* index, so letting a later brush reuse this slot would shift
 * every subsequent SELECTOBJECT index by one.
 */
const RESERVED_SLOT = 'reserved';

/** An object table entry: a usable object, a reserved slot, or a free slot. */
type ObjectSlot = Brush | Pen | GdiFont | typeof RESERVED_SLOT | null;

/** LOGFONT bytes before the facename, and the cap on the facename itself. */
const LOGFONT_FIXED_SIZE = 18;
const FACENAME_MAX = 32;

// ExtTextOut options that prefix the string with a rectangle.
const ETO_OPAQUE = 0x0002;
const ETO_CLIPPED = 0x0004;

interface WmfHeader {
  /** Offset of the first record, in bytes. */
  recordsOffset: number;
  /** Bounding box from the placeable header, in logical units, when present. */
  bbox: { left: number; top: number; right: number; bottom: number } | null;
}

/**
 * Validate and locate the metafile header.
 *
 * A placeable file prefixes the real METAHEADER with 22 bytes carrying the
 * bounding box Office needs to place the drawing; a plain one starts at zero.
 */
function readHeader(view: DataView, length: number): WmfHeader | null {
  let offset = 0;
  let bbox: WmfHeader['bbox'] = null;

  if (length >= PLACEABLE_HEADER_SIZE && view.getUint32(0, true) === PLACEABLE_KEY) {
    bbox = {
      left: view.getInt16(6, true),
      top: view.getInt16(8, true),
      right: view.getInt16(10, true),
      bottom: view.getInt16(12, true),
    };
    offset = PLACEABLE_HEADER_SIZE;
  }

  if (offset + META_HEADER_SIZE > length) return null;

  // Only the fields that identify the format are checked: a wrong Type or
  // HeaderSize means this is not a metafile, while the size and object counts
  // are routinely wrong in files Office still renders.
  const type = view.getUint16(offset, true);
  const headerWords = view.getUint16(offset + 2, true);
  if ((type !== 1 && type !== 2) || headerWords !== META_HEADER_WORDS) return null;

  return { recordsOffset: offset + META_HEADER_SIZE, bbox };
}

/**
 * Parse a WMF file into SVG-ready vector geometry.
 *
 * Returns `null` when the data is not a valid WMF or when no drawable geometry
 * could be recovered, so callers can fall back to their existing behaviour.
 */
export function parseWmfVector(data: Uint8Array): MetafileVectorImage | null {
  if (data.length < META_HEADER_SIZE) return null;

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const header = readHeader(view, data.length);
  if (!header) return null;

  const paths: MetafileVectorImage['paths'] = [];
  const texts: NonNullable<MetafileVectorImage['texts']> = [];
  let dc: DeviceContext = createDeviceContext();
  const dcStack: DeviceContext[] = [];
  /** WMF addresses objects by table index, and a create record takes the lowest free slot. */
  const objects: ObjectSlot[] = [];

  /**
   * DC as it stood when the first geometry was painted; sizes the viewBox.
   *
   * The canvas has to be the coordinate space the emitted path data actually
   * lives in, and that is whatever window was in force when drawing began — not
   * the first window the file happens to declare. Files commonly open with a
   * placeholder 1x1 window and only set the real one just before drawing.
   */
  let frameDc: DeviceContext | null = null;
  /**
   * Whether the file ever set a viewport extent.
   *
   * The scalable map modes derive their scale from viewport/window, but a WMF
   * usually sets only the window: GDI defaults the viewport to the surface the
   * metafile is played back onto, which here is whatever box the picture is
   * drawn into. Scaling by the DC's placeholder 1x1 viewport would divide the
   * whole drawing away, so without a real viewport the mapping stays 1:1 and
   * the logical coordinates are handed to the SVG as-is.
   */
  let viewportExtSet = false;
  /**
   * Whether the file ever set a window extent. The DC's placeholder 1x1 window
   * must not be mistaken for a real one, or a file that sets no window at all
   * would be framed as a single unit instead of falling back to its bounding box.
   */
  let windowExtSet = false;
  let current: Point = { x: 0, y: 0 };
  let emitted = false;

  /**
   * Map a logical point to device space.
   *
   * WMF has no world transform, so this is only the window-to-viewport mapping.
   */
  const scaleOf = (context: DeviceContext): { sx: number; sy: number } =>
    viewportExtSet ? mapScale(context, null) : { sx: 1, sy: 1 };

  const toDeviceWith = (p: Point, context: DeviceContext): Point => {
    const { sx, sy } = scaleOf(context);
    return {
      x: (p.x - context.winOrg.x) * sx + context.vpOrg.x,
      y: (p.y - context.winOrg.y) * sy + context.vpOrg.y,
    };
  };
  const toDevice = (p: Point): Point => toDeviceWith(p, dc);

  /** Record that the file painted something, and freeze the canvas it painted on. */
  const markEmitted = () => {
    if (!frameDc && windowExtSet) frameDc = { ...dc };
    emitted = true;
  };

  /** Paint geometry with the current pen and brush. */
  const draw = (segments: string[], fill: boolean, stroke: boolean) => {
    if (segments.length === 0) return;
    if (paths.length >= MAX_PATHS) return;
    const filled = fill && dc.brush.visible;
    const stroked = stroke && dc.pen.visible;
    if (!filled && !stroked) return;
    const { sx, sy } = scaleOf(dc);
    const penScale = (Math.abs(sx) + Math.abs(sy)) / 2;
    paths.push({
      d: segments.join(' '),
      fill: filled ? dc.brush.color : 'none',
      fillRule: dc.fillRule,
      stroke: stroked ? dc.pen.color : 'none',
      strokeWidth: stroked ? Math.max(dc.pen.width * penScale, 0) : 0,
    });
    markEmitted();
  };

  /**
   * Paint a text run with the current font, colour and alignment.
   *
   * `dx`, when the record carried one, holds the advance of every source *byte*
   * — a double-byte character contributes two entries — and GDI honours it in
   * preference to the font's own metrics. Keeping it is what holds an equation
   * or a positioned label together on a machine without the original typeface.
   */
  const drawText = (bytes: Uint8Array, refX: number, refY: number, dx: number[] | null) => {
    if (texts.length >= MAX_TEXTS) return;
    const font = dc.font;
    const { text, byteCounts } = decodeGdiText(bytes, font.charset, font.facename);
    const chars = Array.from(text);
    if (chars.length === 0) return;

    // Fold each character's bytes together so one advance lines up with one glyph.
    let advances: number[] | null = null;
    if (dx && byteCounts.length === chars.length) {
      advances = [];
      let cursor = 0;
      for (const count of byteCounts) {
        let advance = 0;
        for (let i = 0; i < count && cursor < dx.length; i++, cursor++) advance += dx[cursor];
        advances.push(advance + dc.textCharExtra);
      }
    }

    // TA_UPDATECP makes the DC's current position the reference point and the
    // record's own coordinates dead weight.
    const usesCurrentPoint = (dc.textAlign & TA_UPDATECP) !== 0;
    const originX = usesCurrentPoint ? current.x : refX;
    const originY = usesCurrentPoint ? current.y : refY;

    const { sx, sy } = scaleOf(dc);
    // The window mapping repositions glyphs but never mirrors them.
    const scaleX = Math.abs(sx) || 1;
    const fontSize = font.size * (Math.abs(sy) || 1);
    const origin = toDevice({ x: originX, y: originY });

    const advance = advances
      ? advances.reduce((sum, value) => sum + value, 0)
      : estimateTextAdvance(text, font.size) + dc.textCharExtra * chars.length;
    const anchor = textAnchorOf(dc.textAlign);

    let xs: number[] | undefined;
    if (advances) {
      // Giving every glyph its own position also makes every glyph its own text
      // chunk, so the run's alignment has to be resolved into those positions
      // instead of being left to `text-anchor`.
      const shift = anchor === 'end' ? advance : anchor === 'middle' ? advance / 2 : 0;
      let pen = origin.x - shift * scaleX;
      xs = [];
      for (const step of advances) {
        xs.push(pen);
        pen += step * scaleX;
      }
    }

    texts.push({
      text,
      x: origin.x,
      // The reference point is the top of the cell unless the file says
      // otherwise, so the baseline usually sits an ascent below it.
      y: origin.y + baselineOffset(dc.textAlign, fontSize),
      xs,
      fill: dc.textColor,
      fontFamily: gdiFontStack(font),
      fontSize,
      fontWeight: font.weight,
      italic: font.italic,
      underline: font.underline,
      strikeOut: font.strikeOut,
      // Escapement is counter-clockwise in tenths of a degree; SVG turns the
      // other way.
      rotation: -font.escapement / 10,
      anchor,
    });
    if (usesCurrentPoint) current = { x: originX + advance, y: originY };
    markEmitted();
  };

  /** Put a new object in the lowest free slot, mirroring GDI's allocation. */
  const addObject = (obj: Brush | Pen | GdiFont | typeof RESERVED_SLOT) => {
    const free = objects.indexOf(null);
    if (free >= 0) objects[free] = obj;
    else objects.push(obj);
  };

  let offset = header.recordsOffset;
  let records = 0;

  while (offset + 6 <= data.length && records < MAX_RECORDS) {
    // Record size is a count of 16-bit words, and covers the size and function
    // fields themselves — the smallest legal record is 3 words.
    const sizeWords = view.getUint32(offset, true);
    const func = view.getUint16(offset + 4, true);
    const size = sizeWords * 2;
    if (sizeWords < 3 || offset + size > data.length) break;
    records++;
    if (func === META_EOF) break;

    const body = offset + 6;
    const bodyBytes = size - 6;
    /** Read the n-th 16-bit parameter of the record. */
    const p16 = (index: number): number => view.getInt16(body + index * 2, true);
    const has = (params: number): boolean => bodyBytes >= params * 2;

    switch (func) {
      // --- Coordinate space ---
      // Every record in this group stores its parameters in reverse order, so
      // the y component precedes the x component.
      case META_SETMAPMODE:
        if (has(1)) dc.mapMode = view.getUint16(body, true);
        break;
      case META_SETWINDOWORG:
        if (has(2)) dc.winOrg = { x: p16(1), y: p16(0) };
        break;
      case META_SETWINDOWEXT:
        if (has(2)) {
          dc.winExt = { x: p16(1), y: p16(0) };
          windowExtSet = true;
        }
        break;
      case META_SETVIEWPORTORG:
        if (has(2)) dc.vpOrg = { x: p16(1), y: p16(0) };
        break;
      case META_SETVIEWPORTEXT:
        if (has(2)) {
          dc.vpExt = { x: p16(1), y: p16(0) };
          viewportExtSet = true;
        }
        break;
      case META_OFFSETWINDOWORG:
        if (has(2)) dc.winOrg = { x: dc.winOrg.x + p16(1), y: dc.winOrg.y + p16(0) };
        break;
      case META_OFFSETVIEWPORTORG:
        if (has(2)) dc.vpOrg = { x: dc.vpOrg.x + p16(1), y: dc.vpOrg.y + p16(0) };
        break;
      case META_SCALEWINDOWEXT:
        // yDenom, yNum, xDenom, xNum
        if (has(4) && p16(0) !== 0 && p16(2) !== 0) {
          dc.winExt = {
            x: (dc.winExt.x * p16(3)) / p16(2),
            y: (dc.winExt.y * p16(1)) / p16(0),
          };
        }
        break;
      case META_SCALEVIEWPORTEXT:
        if (has(4) && p16(0) !== 0 && p16(2) !== 0) {
          dc.vpExt = {
            x: (dc.vpExt.x * p16(3)) / p16(2),
            y: (dc.vpExt.y * p16(1)) / p16(0),
          };
          viewportExtSet = true;
        }
        break;

      // --- Device context stack ---
      case META_SAVEDC:
        dcStack.push({ ...dc });
        break;
      case META_RESTOREDC: {
        // A negative level pops that many states; a positive one is an absolute
        // saved-state index, which we approximate by unwinding to it.
        const level = has(1) ? p16(0) : -1;
        const pops =
          level < 0 ? Math.min(-level, dcStack.length) : Math.max(dcStack.length - level, 0);
        for (let i = 0; i < pops; i++) {
          const restored = dcStack.pop();
          if (restored) dc = restored;
        }
        break;
      }

      // --- Fill rule ---
      case META_SETPOLYFILLMODE:
        if (has(1)) {
          dc.fillRule = view.getUint16(body, true) === ALTERNATE_FILL ? 'evenodd' : 'nonzero';
        }
        break;

      // --- Objects ---
      case META_CREATEPENINDIRECT:
        // LogPen: PenStyle(2) Width(POINT16: x,y) ColorRef(4). Only width.x is used.
        if (has(5)) {
          const style = view.getUint16(body, true);
          addObject({
            visible: (style & PS_STYLE_MASK) !== PS_NULL,
            color: colorRefToHex(view.getUint32(body + 6, true)),
            width: Math.abs(p16(1)),
          });
        } else {
          addObject(RESERVED_SLOT);
        }
        break;
      case META_CREATEBRUSHINDIRECT:
        // LogBrush: BrushStyle(2) ColorRef(4) BrushHatch(2)
        if (has(4)) {
          const style = view.getUint16(body, true);
          // A hatched or patterned brush is drawn as its solid foreground
          // colour: closer to the real thing than leaving the shape unfilled.
          addObject({
            visible: style !== BS_NULL,
            color: colorRefToHex(view.getUint32(body + 2, true)),
          });
        } else {
          addObject(RESERVED_SLOT);
        }
        break;
      case META_CREATEFONTINDIRECT:
        addObject(readLogFont(data, view, body, bodyBytes) ?? RESERVED_SLOT);
        break;
      case META_SELECTOBJECT:
        if (has(1)) {
          const index = view.getUint16(body, true);
          // Some producers address stock objects here with the 0x8000 flag set.
          if (index & 0x8000) {
            const stock = stockObject(STOCK_OBJECT_FLAG | (index & 0x7fff));
            if (stock) {
              if (isPen(stock)) dc.pen = stock;
              else dc.brush = stock;
            }
          } else {
            const obj = objects[index];
            if (obj && obj !== RESERVED_SLOT) {
              if (isFont(obj)) dc.font = obj;
              else if (isPen(obj)) dc.pen = obj;
              else dc.brush = obj;
            }
          }
        }
        break;
      case META_DELETEOBJECT:
        if (has(1)) {
          const index = view.getUint16(body, true);
          // Freeing a slot lets the next create record reuse it, as GDI does.
          if (index < objects.length) objects[index] = null;
        }
        break;

      // --- Geometry ---
      case META_MOVETO:
        if (has(2)) current = { x: p16(1), y: p16(0) };
        break;
      case META_LINETO:
        if (has(2)) {
          const to = { x: p16(1), y: p16(0) };
          draw([`M${fmtPoint(toDevice(current))}`, `L${fmtPoint(toDevice(to))}`], false, true);
          current = to;
        }
        break;

      case META_POLYGON:
      case META_POLYLINE: {
        const pts = readPolyPoints(view, body, bodyBytes);
        if (!pts) break;
        const closed = func === META_POLYGON;
        current = pts[pts.length - 1];
        draw(polyToSegments(pts, closed, toDevice), closed, true);
        break;
      }
      case META_POLYPOLYGON: {
        const polygons = readPolyPolyPoints(view, body, bodyBytes);
        if (!polygons) break;
        const segments: string[] = [];
        for (const poly of polygons) {
          if (poly.length === 0) continue;
          segments.push(...polyToSegments(poly, true, toDevice));
          current = poly[poly.length - 1];
        }
        // One path with several subpaths, so the fill rule can carve out holes.
        draw(segments, true, true);
        break;
      }

      case META_RECTANGLE:
      case META_ELLIPSE: {
        // Stored in reverse: bottom, right, top, left.
        if (!has(4)) break;
        const corners = rectCorners(p16(3), p16(2), p16(1), p16(0));
        if (func === META_RECTANGLE) {
          draw(polyToSegments(corners, true, toDevice), true, true);
        } else {
          draw([ellipseSegments(corners, toDevice)], true, true);
        }
        break;
      }
      case META_ROUNDRECT: {
        // Reverse order: ellipse height, ellipse width, bottom, right, top, left.
        if (!has(6)) break;
        const corners = rectCorners(p16(5), p16(4), p16(3), p16(2));
        draw(roundRectSegments(corners, p16(1), p16(0), toDevice), true, true);
        break;
      }

      case META_ARC:
      case META_PIE:
      case META_CHORD: {
        // Reverse order: yEnd, xEnd, yStart, xStart, bottom, right, top, left.
        if (!has(8)) break;
        const corners = rectCorners(p16(7), p16(6), p16(5), p16(4));
        const start = { x: p16(3), y: p16(2) };
        const end = { x: p16(1), y: p16(0) };
        const pts = arcPoints(corners, start, end);
        if (pts.length === 0) break;

        if (func === META_ARC) {
          // An arc is a bare curve: stroked with the pen, never filled.
          draw(polyToSegments(pts, false, toDevice), false, true);
        } else if (func === META_CHORD) {
          draw(polyToSegments(pts, true, toDevice), true, true);
        } else {
          const centre = {
            x: (corners[0].x + corners[2].x) / 2,
            y: (corners[0].y + corners[2].y) / 2,
          };
          draw(polyToSegments([centre, ...pts], true, toDevice), true, true);
        }
        current = pts[pts.length - 1];
        break;
      }

      // --- Text ---
      case META_SETTEXTCOLOR:
        if (has(2)) dc.textColor = colorRefToHex(view.getUint32(body, true));
        break;
      case META_SETTEXTALIGN:
        if (has(1)) dc.textAlign = view.getUint16(body, true);
        break;
      case META_SETTEXTCHAREXTRA:
        if (has(1)) dc.textCharExtra = p16(0);
        break;

      case META_TEXTOUT: {
        // Unusually, the string sits between the length and the position, and
        // it is padded to a whole word before the y, x pair that follows it.
        if (!has(1)) break;
        const count = view.getUint16(body, true);
        const padded = count + (count & 1);
        if (bodyBytes < 2 + padded + 4) break;
        drawText(
          data.subarray(body + 2, body + 2 + count),
          view.getInt16(body + 4 + padded, true),
          view.getInt16(body + 2 + padded, true),
          null,
        );
        break;
      }
      case META_EXTTEXTOUT: {
        // y, x, StringLength, fwOpts, [Rectangle], String, Dx.
        if (!has(4)) break;
        const refY = p16(0);
        const refX = p16(1);
        const count = view.getUint16(body + 4, true);
        const options = view.getUint16(body + 6, true);
        let cursor = body + 8;
        if (options & (ETO_OPAQUE | ETO_CLIPPED)) cursor += 8;
        const padded = count + (count & 1);
        if (cursor + padded > body + bodyBytes) break;

        // The Dx array is optional, and a truncated one is unusable: a partial
        // run of advances would bunch the tail of the string at one point.
        const dxStart = cursor + padded;
        let dx: number[] | null = null;
        if (count > 0 && dxStart + count * 2 <= body + bodyBytes) {
          dx = [];
          for (let i = 0; i < count; i++) dx.push(view.getInt16(dxStart + i * 2, true));
        }
        drawText(data.subarray(cursor, cursor + count), refX, refY, dx);
        break;
      }

      // Recognised but intentionally ignored: they change raster state we do not
      // model, or create an object whose slot still has to be reserved.
      case META_SETBKMODE:
      case META_SETROP2:
      case META_SETSTRETCHBLTMODE:
        break;

      default:
        if (OPAQUE_CREATE_RECORDS.has(func)) addObject(RESERVED_SLOT);
        // Everything else (text, blits, clipping, …) is skipped.
        break;
    }

    offset += size;
  }

  if (!emitted) return null;

  const frame = computeFrame(frameDc, header, toDeviceWith);
  if (!frame) return null;
  return { ...frame, paths, texts };
}

/**
 * Derive the drawing's viewBox in device units.
 *
 * GDI maps the window rectangle onto the viewport, so the window in force while
 * drawing is exactly the canvas the artwork was composed against: mapping its
 * corners through the same transform the geometry went through yields the frame
 * that geometry fills.
 *
 * The placeable header's bounding box is only a fallback. It describes the same
 * picture, but in the header's own physical units — a file whose window is
 * 16000 wide typically declares a bounding box a few thousand units across — so
 * it is the wrong scale for path data expressed in window units, and its origin
 * may be negative where the geometry starts at zero.
 */
function computeFrame(
  frameDc: DeviceContext | null,
  header: WmfHeader,
  toDeviceWith: (p: Point, dc: DeviceContext) => Point,
): Omit<MetafileVectorImage, 'paths'> | null {
  if (frameDc && frameDc.winExt.x !== 0 && frameDc.winExt.y !== 0) {
    const tl = toDeviceWith(frameDc.winOrg, frameDc);
    const br = toDeviceWith(
      { x: frameDc.winOrg.x + frameDc.winExt.x, y: frameDc.winOrg.y + frameDc.winExt.y },
      frameDc,
    );
    const frame = frameFromCorners(tl, br);
    if (frame) return frame;
  }

  const { bbox } = header;
  if (bbox && bbox.right !== bbox.left && bbox.bottom !== bbox.top) {
    const frame = frameFromCorners(
      { x: bbox.left, y: bbox.top },
      { x: bbox.right, y: bbox.bottom },
    );
    if (frame) return frame;
  }

  return null;
}

/** Normalise two opposite corners into a positive-extent frame. */
function frameFromCorners(a: Point, b: Point): Omit<MetafileVectorImage, 'paths'> | null {
  const width = Math.abs(b.x - a.x);
  const height = Math.abs(b.y - a.y);
  if (width <= 0 || height <= 0) return null;
  return { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), width, height };
}

/** Serialize a converted drawing as a standalone SVG document. */
export const wmfVectorToSvg = metafileVectorToSvg;

/**
 * Read the LOGFONT a CreateFontIndirect record carries.
 *
 * The facename is raw bytes in the font's own code page, so a CJK typeface name
 * only survives if it is decoded with the charset recorded beside it.
 */
function readLogFont(
  data: Uint8Array,
  view: DataView,
  body: number,
  bodyBytes: number,
): GdiFont | null {
  if (bodyBytes < LOGFONT_FIXED_SIZE) return null;
  const charset = data[body + 13];
  const nameEnd = Math.min(body + bodyBytes, body + LOGFONT_FIXED_SIZE + FACENAME_MAX);
  const nameBytes = data.subarray(body + LOGFONT_FIXED_SIZE, nameEnd);
  const terminator = nameBytes.indexOf(0);
  return {
    facename: decodeGdiFacename(
      terminator < 0 ? nameBytes : nameBytes.subarray(0, terminator),
      charset,
    ),
    // A zero height asks GDI for a default size; nothing in the file says what
    // it was, so fall back to a legible one rather than to invisible text.
    size: logFontEmSize(view.getInt16(body, true)) || 12,
    weight: view.getInt16(body + 8, true) || 400,
    italic: data[body + 10] !== 0,
    underline: data[body + 11] !== 0,
    strikeOut: data[body + 12] !== 0,
    escapement: view.getInt16(body + 4, true),
    charset,
    pitchAndFamily: data[body + 17],
  };
}

// ---------------------------------------------------------------------------
// Record readers
// ---------------------------------------------------------------------------

/**
 * Read a `Count(2) aPoints[]` payload.
 *
 * Unlike the rectangle records, points here are stored in natural `x, y` order.
 */
function readPolyPoints(view: DataView, body: number, bodyBytes: number): Point[] | null {
  if (bodyBytes < 2) return null;
  const count = view.getUint16(body, true);
  if (count === 0 || count > MAX_POINTS_PER_RECORD) return null;
  if (2 + count * 4 > bodyBytes) return null;
  return readPointArray(view, body + 2, count);
}

/**
 * Read a `NumberOfPolygons(2) aPointsPerPolygon[] aPoints[]` payload.
 *
 * The per-polygon counts are 16-bit here, where EMF uses 32-bit.
 */
function readPolyPolyPoints(view: DataView, body: number, bodyBytes: number): Point[][] | null {
  if (bodyBytes < 4) return null;
  const polyCount = view.getUint16(body, true);
  if (polyCount === 0 || polyCount > 65535) return null;
  if (2 + polyCount * 2 > bodyBytes) return null;

  const countsOffset = body + 2;
  let totalPoints = 0;
  const counts: number[] = [];
  for (let i = 0; i < polyCount; i++) {
    const n = view.getUint16(countsOffset + i * 2, true);
    counts.push(n);
    totalPoints += n;
  }
  if (totalPoints === 0 || totalPoints > MAX_POINTS_PER_RECORD) return null;

  const pointsOffset = countsOffset + polyCount * 2;
  if (pointsOffset - body + totalPoints * 4 > bodyBytes) return null;

  const polygons: Point[][] = [];
  let consumed = 0;
  for (const n of counts) {
    if (n === 0) continue;
    polygons.push(readPointArray(view, pointsOffset + consumed * 4, n));
    consumed += n;
  }
  return polygons.length > 0 ? polygons : null;
}

function readPointArray(view: DataView, offset: number, count: number): Point[] {
  const points: Point[] = new Array(count);
  for (let i = 0; i < count; i++) {
    points[i] = {
      x: view.getInt16(offset + i * 4, true),
      y: view.getInt16(offset + i * 4 + 2, true),
    };
  }
  return points;
}
