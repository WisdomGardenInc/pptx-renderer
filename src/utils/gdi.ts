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

import { cssFontFamilyStack } from '../renderer/fontResolver';

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

// --- Text alignment (SetTextAlign flags) ------------------------------------
/** The record's own reference point is ignored; the DC's current position is used. */
export const TA_UPDATECP = 0x0001;
const TA_HORIZONTAL_MASK = 0x0006;
const TA_RIGHT = 0x0002;
const TA_CENTER = 0x0006;
const TA_VERTICAL_MASK = 0x0018;
const TA_BOTTOM = 0x0008;
const TA_BASELINE = 0x0018;

/**
 * Ascent and descent as a fraction of the em.
 *
 * GDI aligns text with the selected font's real metrics, which are not in the
 * file and not measurable before the SVG is rendered. These are the typical
 * proportions of a text face, and only matter for the top- and bottom-aligned
 * cases: the far more common baseline alignment needs no estimate at all.
 */
const ASCENT_RATIO = 0.88;
const DESCENT_RATIO = 0.12;

/**
 * Character height as a fraction of the LOGFONT cell height.
 *
 * A positive LOGFONT height is the cell height — ascent plus descent — while a
 * CSS font-size is the em. The two differ by roughly this much in a typical
 * text face; a negative height already *is* the em and needs no conversion.
 */
const CELL_HEIGHT_TO_EM = 1 / (ASCENT_RATIO + DESCENT_RATIO);

// --- Safety limits ----------------------------------------------------------
export const MAX_RECORDS = 500_000;
export const MAX_PATHS = 100_000;
export const MAX_POINTS_PER_RECORD = 100_000;
export const MAX_TEXTS = 100_000;

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

/** A single run of text placed on its baseline, in device units. */
export interface MetafileVectorText {
  /** Content, already decoded from the font's charset to Unicode. */
  text: string;
  /** Baseline start in device units, and the origin any rotation turns about. */
  x: number;
  y: number;
  /**
   * Per-character baseline positions in device units, when the record carried a
   * Dx array. GDI spaces glyphs by that array rather than by the font's own
   * advances, so honouring it keeps a run aligned even when the viewer
   * substitutes a different face.
   */
  xs?: number[];
  fill: string;
  /** Ready-to-use CSS font-family list. */
  fontFamily: string;
  /** Em size in device units. */
  fontSize: number;
  fontWeight: number;
  italic: boolean;
  underline: boolean;
  strikeOut: boolean;
  /** Clockwise rotation in degrees about `(x, y)`. */
  rotation: number;
  anchor: 'start' | 'middle' | 'end';
}

/** A converted metafile drawing. Coordinates are device units. */
export interface MetafileVectorImage {
  x: number;
  y: number;
  width: number;
  height: number;
  paths: MetafileVectorPath[];
  /** Text runs, drawn above the geometry in the order GDI emitted them. */
  texts?: MetafileVectorText[];
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

/** A selected logical font, as much of LOGFONT as drawing text needs. */
export interface GdiFont {
  /** Typeface name, already decoded from the charset it was recorded in. */
  facename: string;
  /** Em size in logical units, always positive. */
  size: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeOut: boolean;
  /** Baseline rotation in tenths of a degree, counter-clockwise. */
  escapement: number;
  /** LOGFONT CharSet, which decides how the text bytes are decoded. */
  charset: number;
  /** LOGFONT PitchAndFamily; only its family bits are used, to pick a fallback. */
  pitchAndFamily: number;
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
  font: GdiFont;
  textColor: string;
  /** SetTextAlign flags. */
  textAlign: number;
  /** SetTextCharacterExtra: extra space after every character, in logical units. */
  textCharExtra: number;
  fillRule: 'nonzero' | 'evenodd';
  mapMode: number;
  winOrg: Point;
  winExt: Point;
  vpOrg: Point;
  vpExt: Point;
  xform: Xform;
}

/** The system font a DC starts with, standing in for GDI's SYSTEM_FONT. */
const DEFAULT_FONT: GdiFont = {
  facename: '',
  size: 12,
  weight: 400,
  italic: false,
  underline: false,
  strikeOut: false,
  escapement: 0,
  charset: 0,
  pitchAndFamily: 0,
};

/** A pen, brush and font a freshly created DC starts with. */
export function createDeviceContext(): DeviceContext {
  return {
    brush: { visible: true, color: '#FFFFFF' },
    pen: { visible: true, color: '#000000', width: 0 },
    font: DEFAULT_FONT,
    textColor: '#000000',
    textAlign: 0,
    textCharExtra: 0,
    fillRule: 'evenodd',
    mapMode: MM_TEXT,
    winOrg: { x: 0, y: 0 },
    winExt: { x: 1, y: 1 },
    vpOrg: { x: 0, y: 0 },
    vpExt: { x: 1, y: 1 },
    xform: IDENTITY,
  };
}

/** Distinguishes the object kinds stored in one GDI object table. */
export function isPen(obj: Brush | Pen | GdiFont): obj is Pen {
  return 'width' in obj;
}

export function isFont(obj: Brush | Pen | GdiFont): obj is GdiFont {
  return 'facename' in obj;
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
// Text
// ---------------------------------------------------------------------------

const SYMBOL_CHARSET = 2;

/**
 * Windows code page for each LOGFONT CharSet, as a `TextDecoder` label.
 *
 * A metafile stores text as raw bytes in the selected font's charset, with no
 * marker of its own, so the font is the only thing that says how to read them —
 * get this wrong on a CJK file and every glyph is mojibake.
 */
const CHARSET_ENCODINGS: Record<number, string> = {
  0: 'windows-1252', // ANSI
  128: 'shift_jis',
  129: 'euc-kr',
  130: 'euc-kr', // Johab, close enough for face names
  134: 'gbk', // GB2312
  136: 'big5',
  161: 'windows-1253', // Greek
  162: 'windows-1254', // Turkish
  163: 'windows-1258', // Vietnamese
  177: 'windows-1255', // Hebrew
  178: 'windows-1256', // Arabic
  186: 'windows-1257', // Baltic
  204: 'windows-1251', // Cyrillic
  222: 'windows-874', // Thai
  238: 'windows-1250', // Central European
  255: 'windows-1252', // OEM
};

/**
 * Adobe Symbol encoding, from 0x20 up, as Unicode.
 *
 * The Symbol font is not text in any code page: its bytes are glyph slots, and
 * a run reading `ec ef ef ee` is the four pieces of a tall curly brace, not
 * accented Latin letters. Equation editors emit almost nothing else, so mapping
 * the slots to their Unicode equivalents is what makes those files legible.
 */
const SYMBOL_TO_UNICODE =
  ' !\u2200#\u2203%&\u220b()\u2217+,\u2212./0123456789:;<=>?' +
  '\u2245\u0391\u0392\u03a7\u0394\u0395\u03a6\u0393\u0397\u0399\u03d1\u039a\u039b\u039c\u039d\u039f' +
  '\u03a0\u0398\u03a1\u03a3\u03a4\u03a5\u03c2\u03a9\u039e\u03a8\u0396[\u2234]\u22a5_' +
  '\u203e\u03b1\u03b2\u03c7\u03b4\u03b5\u03c6\u03b3\u03b7\u03b9\u03d5\u03ba\u03bb\u03bc\u03bd\u03bf' +
  '\u03c0\u03b8\u03c1\u03c3\u03c4\u03c5\u03d6\u03c9\u03be\u03c8\u03b6{|}\u223c' +
  '                                 ' + // 0x7f-0x9f: unassigned
  '\u20ac\u03d2\u2032\u2264\u2044\u221e\u0192\u2663\u2666\u2665\u2660\u2194\u2190\u2191\u2192\u2193' +
  '\u00b0\u00b1\u2033\u2265\u00d7\u221d\u2202\u2022\u00f7\u2260\u2261\u2248\u2026\u23d0\u2500\u21b5' +
  '\u2135\u2111\u211c\u2118\u2297\u2295\u2205\u2229\u222a\u2283\u2287\u2284\u2282\u2286\u2208\u2209' +
  '\u2220\u2207\u00ae\u00a9\u2122\u220f\u221a\u22c5\u00ac\u2227\u2228\u21d4\u21d0\u21d1\u21d2\u21d3' +
  '\u25ca\u2329\u00ae\u00a9\u2122\u2211\u239b\u239c\u239d\u23a1\u23a2\u23a3\u23a7\u23a8\u23a9\u23aa' +
  ' \u232a\u222b\u2320\u23ae\u2321\u239e\u239f\u23a0\u23a4\u23a5\u23a6\u23ab\u23ac\u23ad';

const decoderCache = new Map<string, TextDecoder | null>();

function decoderFor(label: string): TextDecoder | null {
  const cached = decoderCache.get(label);
  if (cached !== undefined) return cached;
  let decoder: TextDecoder | null;
  try {
    decoder = new TextDecoder(label);
  } catch {
    decoder = null; // Runtime without that legacy code page.
  }
  decoderCache.set(label, decoder);
  return decoder;
}

function isSymbolFont(facename: string): boolean {
  return facename.trim().toLowerCase() === 'symbol';
}

/** Charsets whose code pages are multi-byte, where one byte is not one glyph. */
const DBCS_CHARSETS = new Set([128, 129, 130, 134, 136]);

interface DecodedGdiText {
  text: string;
  /**
   * Source bytes each character was decoded from.
   *
   * A metafile's Dx array is indexed by byte, not by character, so a run of
   * double-byte text needs this to fold two spacing entries onto one glyph.
   */
  byteCounts: number[];
}

/** Every character came from exactly one byte. */
function singleByteRuns(text: string): DecodedGdiText {
  return { text, byteCounts: Array.from(text, () => 1) };
}

/**
 * Decode metafile text bytes using the charset the font was created with.
 *
 * Symbol-charset fonts other than Symbol itself (Wingdings and friends) keep
 * their bytes, lifted into the private-use area their own `cmap` uses, so a
 * viewer that has the face installed still draws the intended dingbat.
 */
export function decodeGdiText(
  bytes: Uint8Array,
  charset: number,
  facename: string,
): DecodedGdiText {
  if (isSymbolFont(facename)) {
    let out = '';
    for (const byte of bytes) {
      const mapped = byte >= 0x20 ? SYMBOL_TO_UNICODE[byte - 0x20] : undefined;
      out += mapped ?? String.fromCharCode(byte);
    }
    return singleByteRuns(out);
  }
  if (charset === SYMBOL_CHARSET) {
    let out = '';
    for (const byte of bytes) out += String.fromCharCode(0xf000 + byte);
    return singleByteRuns(out);
  }

  const label = CHARSET_ENCODINGS[charset] ?? 'windows-1252';
  if (!DBCS_CHARSETS.has(charset)) {
    const decoder = decoderFor(label);
    if (decoder) return singleByteRuns(decoder.decode(bytes));
    let out = '';
    for (const byte of bytes) out += String.fromCharCode(byte);
    return singleByteRuns(out);
  }
  return decodeDbcs(bytes, label);
}

/**
 * Decode a multi-byte run one byte at a time, recording how many bytes each
 * character consumed. Feeding the decoder in single-byte chunks is what makes
 * that boundary observable: it emits nothing until a character is complete.
 */
function decodeDbcs(bytes: Uint8Array, label: string): DecodedGdiText {
  let decoder: TextDecoder;
  try {
    decoder = new TextDecoder(label);
  } catch {
    let out = '';
    for (const byte of bytes) out += String.fromCharCode(byte);
    return singleByteRuns(out);
  }

  let text = '';
  const byteCounts: number[] = [];
  let pending = 0;
  const take = (piece: string) => {
    for (const _ of piece) {
      byteCounts.push(pending);
      pending = 0;
    }
    text += piece;
  };
  for (let i = 0; i < bytes.length; i++) {
    pending++;
    take(decoder.decode(bytes.subarray(i, i + 1), { stream: true }));
  }
  take(decoder.decode());
  return { text, byteCounts };
}

/**
 * Decode a LOGFONT facename.
 *
 * The name is ordinary text in the font's code page even when the font's
 * charset marks its *contents* as symbol slots, so the symbol mappings that
 * apply to the text drawn with it must not apply to its name.
 */
export function decodeGdiFacename(bytes: Uint8Array, charset: number): string {
  return decodeGdiText(bytes, charset === SYMBOL_CHARSET ? 0 : charset, '').text;
}

/** Normalise a LOGFONT height to a positive em size. */
export function logFontEmSize(height: number): number {
  return height < 0 ? -height : height * CELL_HEIGHT_TO_EM;
}

/** Generic CSS family implied by the LOGFONT PitchAndFamily family bits. */
function genericFamily(pitchAndFamily: number): string {
  switch (pitchAndFamily & 0xf0) {
    case 0x10:
      return 'serif'; // FF_ROMAN
    case 0x30:
      return 'monospace'; // FF_MODERN
    case 0x40:
      return 'cursive'; // FF_SCRIPT
    case 0x50:
      return 'fantasy'; // FF_DECORATIVE
    default:
      return 'sans-serif'; // FF_SWISS and FF_DONTCARE
  }
}

/**
 * CSS font stack for a logical font.
 *
 * Symbol runs are decoded to real Unicode rather than kept as Symbol slots, so
 * they ask for a maths-capable face instead: the brace and bracket pieces an
 * equation is built from live outside any text font's repertoire.
 */
export function gdiFontStack(font: GdiFont): string {
  if (isSymbolFont(font.facename)) {
    return cssFontFamilyStack([
      'Cambria Math',
      'STIX Two Math',
      'Segoe UI Symbol',
      'Apple Symbols',
      'serif',
    ]);
  }
  const generic = genericFamily(font.pitchAndFamily);
  if (!font.facename) return generic;
  // The generic goes last: a CJK face brings its own substitution chain, and
  // putting a bare `sans-serif` ahead of it would win before the chain is tried.
  return `${cssFontFamilyStack([font.facename])}, ${generic}`;
}

/**
 * Advance of a run in em units, used only to place text the file did not give a
 * Dx array for. A CJK or fullwidth character occupies a full em, most others a
 * half — the same rule GDI's own fallback metrics approximate.
 */
export function estimateTextAdvance(text: string, emSize: number): number {
  let ems = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    ems += isWideCodePoint(code) ? 1 : 0.5;
  }
  return ems * emSize;
}

function isWideCodePoint(code: number): boolean {
  return (
    (code >= 0x1100 && code <= 0x115f) || // Hangul Jamo
    (code >= 0x2e80 && code <= 0xa4cf) || // CJK radicals through Yi
    (code >= 0xac00 && code <= 0xd7a3) || // Hangul syllables
    (code >= 0xf900 && code <= 0xfaff) || // CJK compatibility ideographs
    (code >= 0xfe30 && code <= 0xfe4f) || // CJK compatibility forms
    (code >= 0xff00 && code <= 0xff60) || // Fullwidth forms
    (code >= 0xffe0 && code <= 0xffe6) ||
    (code >= 0x20000 && code <= 0x3fffd) // CJK extensions B and beyond
  );
}

/** Where the horizontal reference point sits within the run. */
export function textAnchorOf(textAlign: number): 'start' | 'middle' | 'end' {
  switch (textAlign & TA_HORIZONTAL_MASK) {
    case TA_RIGHT:
      return 'end';
    case TA_CENTER:
      return 'middle';
    default:
      return 'start';
  }
}

/**
 * Distance from the reference point down to the baseline, in logical units.
 *
 * GDI's default is TA_TOP, where the reference point is the top of the cell;
 * the baseline sits an ascent below it.
 */
export function baselineOffset(textAlign: number, emSize: number): number {
  switch (textAlign & TA_VERTICAL_MASK) {
    case TA_BASELINE:
      return 0;
    case TA_BOTTOM:
      return -DESCENT_RATIO * emSize;
    default:
      return ASCENT_RATIO * emSize;
  }
}

// ---------------------------------------------------------------------------
// SVG serialization
// ---------------------------------------------------------------------------

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Serialize one text run as an SVG `<text>` element. */
function textToSvg(t: MetafileVectorText): string {
  // A multi-value x places every glyph itself, which also makes each glyph its
  // own text chunk — so the anchor has to have been folded into the positions
  // already, and only the single-position form can still carry one.
  const x = t.xs && t.xs.length > 0 ? t.xs.map(fmtNumber).join(' ') : fmtNumber(t.x);
  const attrs = [`x="${x}"`, `y="${fmtNumber(t.y)}"`, `fill="${t.fill}"`];
  attrs.push(`font-family="${escapeXml(t.fontFamily)}"`);
  attrs.push(`font-size="${fmtNumber(t.fontSize)}"`);
  if (t.fontWeight !== 400) attrs.push(`font-weight="${Math.round(t.fontWeight)}"`);
  if (t.italic) attrs.push('font-style="italic"');
  const decorations: string[] = [];
  if (t.underline) decorations.push('underline');
  if (t.strikeOut) decorations.push('line-through');
  if (decorations.length > 0) attrs.push(`text-decoration="${decorations.join(' ')}"`);
  if (!t.xs && t.anchor !== 'start') attrs.push(`text-anchor="${t.anchor}"`);
  if (t.rotation !== 0) {
    attrs.push(`transform="rotate(${fmtNumber(t.rotation)} ${fmtNumber(t.x)} ${fmtNumber(t.y)})"`);
  }
  attrs.push('xml:space="preserve"');
  return `<text ${attrs.join(' ')}>${escapeXml(t.text)}</text>`;
}

/** Serialize a converted drawing as a standalone SVG document. */
export function metafileVectorToSvg(image: MetafileVectorImage): string {
  const body =
    image.paths
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
      .join('') + (image.texts ?? []).map(textToSvg).join('');
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
