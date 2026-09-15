import { describe, it, expect } from 'vitest';
import { parseWmfVector, wmfVectorToSvg } from '../../../src/utils/wmfVector';

// ---------------------------------------------------------------------------
// Helpers to build synthetic WMF binary data
// ---------------------------------------------------------------------------

const PLACEABLE_KEY = 0x9ac6cdd7;

/**
 * Build a record: a 4-byte word count, a 2-byte function code, then 16-bit
 * parameters. The size field counts 16-bit words and includes itself.
 */
function record(func: number, params: number[] = []): Uint8Array {
  const size = 6 + params.length * 2;
  const buf = new Uint8Array(size);
  const view = new DataView(buf.buffer);
  view.setUint32(0, size / 2, true);
  view.setUint16(4, func, true);
  params.forEach((p, i) => view.setInt16(6 + i * 2, p, true));
  return buf;
}

/** A record whose parameters are raw bytes rather than 16-bit words. */
function rawRecord(func: number, body: Uint8Array): Uint8Array {
  const size = 6 + body.length;
  const buf = new Uint8Array(size);
  const view = new DataView(buf.buffer);
  view.setUint32(0, size / 2, true);
  view.setUint16(4, func, true);
  buf.set(body, 6);
  return buf;
}

/** 18-byte METAHEADER. */
function metaHeader(): Uint8Array {
  const buf = new Uint8Array(18);
  const view = new DataView(buf.buffer);
  view.setUint16(0, 1, true); // Type: memory metafile
  view.setUint16(2, 9, true); // HeaderSize, in words
  view.setUint16(4, 0x0300, true); // Version: Windows 3.0
  view.setUint16(12, 4, true); // NumberOfObjects
  return buf;
}

/** 22-byte Aldus placeable header. */
function placeableHeader(
  bbox: [number, number, number, number] = [0, 0, 200, 100],
  inch = 96,
): Uint8Array {
  const buf = new Uint8Array(22);
  const view = new DataView(buf.buffer);
  view.setUint32(0, PLACEABLE_KEY, true);
  bbox.forEach((v, i) => view.setInt16(6 + i * 2, v, true));
  view.setUint16(14, inch, true);
  return buf;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, p) => sum + p.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

const META_EOF = 0x0000;
const META_SETPOLYFILLMODE = 0x0106;
const META_SAVEDC = 0x001e;
const META_RESTOREDC = 0x0127;
const META_SELECTOBJECT = 0x012d;
const META_DELETEOBJECT = 0x01f0;
const META_SETWINDOWORG = 0x020b;
const META_SETWINDOWEXT = 0x020c;
const META_LINETO = 0x0213;
const META_MOVETO = 0x0214;
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
const META_SETTEXTALIGN = 0x012e;
const META_SETTEXTCOLOR = 0x0209;
const META_TEXTOUT = 0x0521;
const META_EXTTEXTOUT = 0x0a32;
const META_DIBBITBLT = 0x0940;

/** A window covering 0,0..200,100, which makes device units equal logical units. */
const windowRecords = [
  record(META_SETWINDOWORG, [0, 0]),
  record(META_SETWINDOWEXT, [100, 200]), // reverse order: height, width
];

function buildWmf(records: Uint8Array[], options: { placeable?: boolean } = {}): Uint8Array {
  const { placeable = true } = options;
  return concat([
    ...(placeable ? [placeableHeader()] : []),
    metaHeader(),
    ...records,
    record(META_EOF),
  ]);
}

/** Build a solid brush record body: BrushStyle(2) ColorRef(4) BrushHatch(2). */
function brushBody(style: number, color: number): Uint8Array {
  const buf = new Uint8Array(8);
  const view = new DataView(buf.buffer);
  view.setUint16(0, style, true);
  view.setUint32(2, color, true);
  return buf;
}

/** Build a pen record body: PenStyle(2) Width(x,y) ColorRef(4). */
function penBody(style: number, width: number, color: number): Uint8Array {
  const buf = new Uint8Array(10);
  const view = new DataView(buf.buffer);
  view.setUint16(0, style, true);
  view.setInt16(2, width, true);
  view.setUint32(6, color, true);
  return buf;
}

/** Build a LOGFONT body: 18 fixed bytes, then a NUL-terminated facename. */
function fontBody(
  options: {
    height?: number;
    escapement?: number;
    weight?: number;
    italic?: boolean;
    underline?: boolean;
    charset?: number;
    facename?: string;
  } = {},
): Uint8Array {
  const name = options.facename ?? '';
  // Records are counted in whole words, so the padded name keeps the body even.
  const buf = new Uint8Array(18 + name.length + (name.length % 2 === 0 ? 2 : 1));
  const view = new DataView(buf.buffer);
  view.setInt16(0, options.height ?? -20, true);
  view.setInt16(4, options.escapement ?? 0, true);
  view.setInt16(8, options.weight ?? 400, true);
  buf[10] = options.italic ? 1 : 0;
  buf[11] = options.underline ? 1 : 0;
  buf[13] = options.charset ?? 0;
  for (let i = 0; i < name.length; i++) buf[18 + i] = name.charCodeAt(i);
  return buf;
}

/**
 * Build an ExtTextOut body: y, x, StringLength, fwOpts, [Rectangle], String
 * (padded to a word), Dx. The rectangle is only present for the opaque and
 * clipped options.
 */
function extTextOutBody(
  x: number,
  y: number,
  text: Uint8Array,
  dx?: number[],
  options = 0,
): Uint8Array {
  const rect = options & 0x0006 ? 8 : 0;
  const padded = text.length + (text.length & 1);
  const buf = new Uint8Array(8 + rect + padded + (dx ? dx.length * 2 : 0));
  const view = new DataView(buf.buffer);
  view.setInt16(0, y, true);
  view.setInt16(2, x, true);
  view.setUint16(4, text.length, true);
  view.setUint16(6, options, true);
  buf.set(text, 8 + rect);
  dx?.forEach((value, i) => view.setInt16(8 + rect + padded + i * 2, value, true));
  return buf;
}

/** Build a TextOut body: StringLength, String (padded), then y, x. */
function textOutBody(x: number, y: number, text: Uint8Array): Uint8Array {
  const padded = text.length + (text.length & 1);
  const buf = new Uint8Array(2 + padded + 4);
  const view = new DataView(buf.buffer);
  view.setUint16(0, text.length, true);
  buf.set(text, 2);
  view.setInt16(2 + padded, y, true);
  view.setInt16(4 + padded, x, true);
  return buf;
}

function ascii(text: string): Uint8Array {
  return Uint8Array.from(text, (ch) => ch.charCodeAt(0));
}

/** Pack a `Count(2) aPoints[]` payload, points in natural x,y order. */
function polyBody(pts: Array<[number, number]>): Uint8Array {
  const buf = new Uint8Array(2 + pts.length * 4);
  const view = new DataView(buf.buffer);
  view.setUint16(0, pts.length, true);
  pts.forEach(([x, y], i) => {
    view.setInt16(2 + i * 4, x, true);
    view.setInt16(2 + i * 4 + 2, y, true);
  });
  return buf;
}

// ---------------------------------------------------------------------------
// Header handling
// ---------------------------------------------------------------------------

describe('parseWmfVector — headers', () => {
  it('parses a placeable file, skipping the 22-byte Aldus prefix', () => {
    const wmf = buildWmf([
      ...windowRecords,
      rawRecord(META_POLYGON, polyBody([[10, 10], [60, 10], [60, 40]])),
    ]);
    const image = parseWmfVector(wmf);
    expect(image).not.toBeNull();
    expect(image!.paths).toHaveLength(1);
  });

  it('parses a plain file that has no placeable header', () => {
    const wmf = buildWmf(
      [...windowRecords, rawRecord(META_POLYGON, polyBody([[0, 0], [10, 0], [10, 10]]))],
      { placeable: false },
    );
    expect(parseWmfVector(wmf)).not.toBeNull();
  });

  it('rejects data whose METAHEADER is not a metafile header', () => {
    const wmf = buildWmf([...windowRecords]);
    // HeaderSize must be 9 words; corrupt it in place (after the placeable prefix).
    new DataView(wmf.buffer).setUint16(22 + 2, 7, true);
    expect(parseWmfVector(wmf)).toBeNull();
  });

  it('rejects truncated data', () => {
    expect(parseWmfVector(new Uint8Array(8))).toBeNull();
  });

  it('returns null when no record draws anything', () => {
    expect(parseWmfVector(buildWmf([...windowRecords]))).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// The viewBox
// ---------------------------------------------------------------------------

describe('parseWmfVector — frame', () => {
  it('derives the viewBox from the window rectangle', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, rawRecord(META_POLYGON, polyBody([[0, 0], [10, 10]]))]),
    );
    expect(image).toMatchObject({ x: 0, y: 0, width: 200, height: 100 });
  });

  it('honours a window origin offset so geometry lands inside the viewBox', () => {
    const image = parseWmfVector(
      buildWmf([
        record(META_SETWINDOWORG, [50, 100]), // reverse order: y, x
        record(META_SETWINDOWEXT, [100, 200]),
        rawRecord(META_POLYGON, polyBody([[100, 50], [300, 50], [300, 150]])),
      ]),
    );
    // The origin is subtracted, so the drawing starts at the viewBox origin.
    expect(image).toMatchObject({ x: 0, y: 0, width: 200, height: 100 });
    expect(image!.paths[0].d).toContain('M0,0');
  });

  it('frames by the window in force when drawing starts, not the first one declared', () => {
    // Real clip art routinely opens with a placeholder 1x1 window and sets the
    // real one just before drawing. Framing by the first window collapses the
    // whole picture into a single unit.
    const image = parseWmfVector(
      buildWmf([
        record(META_SETWINDOWORG, [0, 0]),
        record(META_SETWINDOWEXT, [1, 1]),
        record(META_SETWINDOWORG, [0, 0]),
        record(META_SETWINDOWEXT, [12000, 16000]), // height, width
        rawRecord(META_POLYGON, polyBody([[0, 0], [16000, 0], [16000, 12000]])),
      ]),
    );
    expect(image).toMatchObject({ width: 16000, height: 12000 });
  });

  it('ignores the placeable bounding box when a window defines the canvas', () => {
    // The bounding box is in the header's own physical units, so it is the wrong
    // scale for path data expressed in window units.
    const wmf = concat([
      placeableHeader([0, 0, 3961, 3370]), // a few thousand units across
      metaHeader(),
      record(META_SETWINDOWORG, [0, 0]),
      record(META_SETWINDOWEXT, [13615, 16000]),
      rawRecord(META_POLYGON, polyBody([[0, 0], [16000, 13615]])),
      record(META_EOF),
    ]);
    expect(parseWmfVector(wmf)).toMatchObject({ width: 16000, height: 13615 });
  });

  it('does not scale by the viewport when the file never sets one', () => {
    // MM_ANISOTROPIC takes its scale from viewport/window, but a file that sets
    // only the window relies on the player supplying the viewport. Using the
    // DC's placeholder 1x1 viewport would divide the drawing away.
    const image = parseWmfVector(
      buildWmf([
        record(0x0103, [8]), // SETMAPMODE MM_ANISOTROPIC
        record(META_SETWINDOWORG, [0, 0]),
        record(META_SETWINDOWEXT, [12000, 16000]),
        rawRecord(META_POLYGON, polyBody([[0, 0], [16000, 12000]])),
      ]),
    );
    expect(image).toMatchObject({ width: 16000, height: 12000 });
    expect(image!.paths[0].d).toBe('M0,0 L16000,12000 Z');
  });

  it('honours a viewport the file does set', () => {
    const image = parseWmfVector(
      buildWmf([
        record(0x0103, [8]), // SETMAPMODE MM_ANISOTROPIC
        record(META_SETWINDOWORG, [0, 0]),
        record(META_SETWINDOWEXT, [1000, 2000]), // height 1000, width 2000
        record(0x020d, [0, 0]), // SETVIEWPORTORG
        record(0x020e, [500, 1000]), // SETVIEWPORTEXT height 500, width 1000
        rawRecord(META_POLYGON, polyBody([[0, 0], [2000, 1000]])),
      ]),
    );
    // A half-size viewport halves both the geometry and the frame.
    expect(image).toMatchObject({ width: 1000, height: 500 });
    expect(image!.paths[0].d).toBe('M0,0 L1000,500 Z');
  });

  it('falls back to the placeable bounding box without a window extent', () => {
    const wmf = concat([
      placeableHeader([0, 0, 640, 480]),
      metaHeader(),
      rawRecord(META_POLYGON, polyBody([[0, 0], [10, 0], [10, 10]])),
      record(META_EOF),
    ]);
    expect(parseWmfVector(wmf)).toMatchObject({ x: 0, y: 0, width: 640, height: 480 });
  });
});

// ---------------------------------------------------------------------------
// Reverse-ordered parameters
//
// These records store their coordinates backwards. Getting the order wrong still
// produces a plausible-looking shape, so each is pinned to exact path data.
// ---------------------------------------------------------------------------

describe('parseWmfVector — reverse-ordered coordinates', () => {
  it('reads META_RECTANGLE as bottom, right, top, left', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, record(META_RECTANGLE, [40, 60, 10, 20])]),
    );
    // left=20 top=10 right=60 bottom=40
    expect(image!.paths[0].d).toBe('M20,10 L60,10 L60,40 L20,40 Z');
  });

  it('reads META_SETWINDOWEXT as height, width', () => {
    const image = parseWmfVector(
      buildWmf([
        record(META_SETWINDOWORG, [0, 0]),
        record(META_SETWINDOWEXT, [300, 400]), // height=300 width=400
        rawRecord(META_POLYGON, polyBody([[0, 0], [1, 1]])),
      ]),
    );
    expect(image).toMatchObject({ width: 400, height: 300 });
  });

  it('reads META_MOVETO and META_LINETO as y, x', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_MOVETO, [30, 10]), // y=30 x=10
        record(META_LINETO, [80, 70]), // y=80 x=70
      ]),
    );
    expect(image!.paths[0].d).toBe('M10,30 L70,80');
  });

  it('reads META_ELLIPSE as bottom, right, top, left', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, record(META_ELLIPSE, [100, 200, 0, 0])]),
    );
    // A 200x100 ellipse starts at the left edge of its vertical centre.
    expect(image!.paths[0].d).toMatch(/^M0,50 /);
  });

  it('reads META_ROUNDRECT as ellipse height, ellipse width, then the rect', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        // bottom=100 right=200 top=0 left=0, ellipse width=40 height=20
        record(META_ROUNDRECT, [20, 40, 100, 200, 0, 0]),
      ]),
    );
    // Corner radii are half the ellipse axes: 20 across, 10 down.
    expect(image!.paths[0].d).toMatch(/^M20,0 L180,0/);
    expect(image!.paths[0].d).toContain('L200,90');
  });

  it('keeps poly* points in natural x,y order, unlike the rectangle records', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, rawRecord(META_POLYGON, polyBody([[10, 20], [30, 40]]))]),
    );
    expect(image!.paths[0].d).toBe('M10,20 L30,40 Z');
  });
});

// ---------------------------------------------------------------------------
// Geometry records
// ---------------------------------------------------------------------------

describe('parseWmfVector — geometry', () => {
  it('closes and fills a polygon but leaves a polyline open and unfilled', () => {
    const pts: Array<[number, number]> = [[0, 0], [50, 0], [50, 50]];
    const polygon = parseWmfVector(buildWmf([...windowRecords, rawRecord(META_POLYGON, polyBody(pts))]));
    const polyline = parseWmfVector(buildWmf([...windowRecords, rawRecord(META_POLYLINE, polyBody(pts))]));

    expect(polygon!.paths[0].d.endsWith(' Z')).toBe(true);
    expect(polygon!.paths[0].fill).not.toBe('none');
    expect(polyline!.paths[0].d.endsWith(' Z')).toBe(false);
    expect(polyline!.paths[0].fill).toBe('none');
  });

  it('emits polypolygon contours as one path so the fill rule can carve holes', () => {
    const body = new Uint8Array(2 + 2 * 2 + 8 * 4);
    const view = new DataView(body.buffer);
    view.setUint16(0, 2, true); // two polygons
    view.setUint16(2, 4, true); // outer: 4 points
    view.setUint16(4, 4, true); // inner: 4 points
    const pts: Array<[number, number]> = [
      [0, 0], [100, 0], [100, 100], [0, 100],
      [25, 25], [75, 25], [75, 75], [25, 75],
    ];
    pts.forEach(([x, y], i) => {
      view.setInt16(6 + i * 4, x, true);
      view.setInt16(6 + i * 4 + 2, y, true);
    });

    const image = parseWmfVector(
      buildWmf([...windowRecords, record(META_SETPOLYFILLMODE, [1]), rawRecord(META_POLYPOLYGON, body)]),
    );
    expect(image!.paths).toHaveLength(1);
    expect(image!.paths[0].d.match(/M/g)).toHaveLength(2);
    expect(image!.paths[0].fillRule).toBe('evenodd');
  });

  it('strokes an arc without filling it, but fills a pie and a chord', () => {
    const arcParams = [50, 100, 0, 200, 100, 200, 0, 0];
    const arc = parseWmfVector(buildWmf([...windowRecords, record(META_ARC, arcParams)]));
    const pie = parseWmfVector(buildWmf([...windowRecords, record(META_PIE, arcParams)]));
    const chord = parseWmfVector(buildWmf([...windowRecords, record(META_CHORD, arcParams)]));

    expect(arc!.paths[0].fill).toBe('none');
    expect(arc!.paths[0].d.endsWith(' Z')).toBe(false);
    expect(pie!.paths[0].fill).not.toBe('none');
    expect(chord!.paths[0].fill).not.toBe('none');
    // A pie is drawn from the centre of its bounding box; a chord is not.
    expect(pie!.paths[0].d.startsWith('M100,50')).toBe(true);
    expect(chord!.paths[0].d.startsWith('M100,50')).toBe(false);
  });

  it('sweeps an arc counter-clockwise, as GDI does', () => {
    // A 200x100 ellipse at the origin, from the 3 o'clock ray to the 12 o'clock ray.
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_ARC, [0, 100, 50, 200, 100, 200, 0, 0]),
      ]),
    );
    const d = image!.paths[0].d;
    // Counter-clockwise runs 3 o'clock -> 12 o'clock the short way, so the sweep
    // stays in the upper right quadrant instead of going the long way round the
    // bottom. Sampling the midpoint of that quadrant pins the direction.
    expect(d.startsWith('M200,50')).toBe(true);
    expect(d.endsWith('L100,0')).toBe(true);
    const ys = [...d.matchAll(/,(-?[\d.]+)/g)].map((m) => Number(m[1]));
    expect(Math.max(...ys)).toBe(50); // never dips below the vertical centre
  });
});

// ---------------------------------------------------------------------------
// The object table
// ---------------------------------------------------------------------------

describe('parseWmfVector — object table', () => {
  it('selects a brush and pen by table index', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x0000ff)), // red, as 0x00bbggrr
        rawRecord(META_CREATEPENINDIRECT, penBody(0, 3, 0xff0000)), // blue
        record(META_SELECTOBJECT, [0]),
        record(META_SELECTOBJECT, [1]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#FF0000');
    expect(image!.paths[0].stroke).toBe('#0000FF');
    expect(image!.paths[0].strokeWidth).toBe(3);
  });

  it('reserves a slot for create records it cannot use, keeping later indices right', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        // A font takes slot 0 even though it draws nothing.
        rawRecord(META_CREATEFONTINDIRECT, new Uint8Array(18)),
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x00ff00)), // green, slot 1
        record(META_SELECTOBJECT, [1]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#00FF00');
  });

  it('lets a deleted slot be reused by the next create record', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x0000ff)), // slot 0
        record(META_DELETEOBJECT, [0]),
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x00ff00)), // reuses slot 0
        record(META_SELECTOBJECT, [0]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#00FF00');
  });

  it('treats a BS_NULL brush as hollow and a PS_NULL pen as invisible', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(1, 0x0000ff)), // BS_NULL
        record(META_SELECTOBJECT, [0]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('none');
    expect(image!.paths[0].stroke).not.toBe('none');
  });

  it('drops a shape whose pen and brush are both invisible', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(1, 0)), // BS_NULL
        rawRecord(META_CREATEPENINDIRECT, penBody(5, 1, 0)), // PS_NULL
        record(META_SELECTOBJECT, [0]),
        record(META_SELECTOBJECT, [1]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image).toBeNull();
  });

  it('ignores a select of an empty or out-of-range slot', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, record(META_SELECTOBJECT, [7]), record(META_RECTANGLE, [50, 50, 0, 0])]),
    );
    expect(image!.paths[0].fill).toBe('#FFFFFF'); // the default brush
  });

  it('resolves a stock object addressed with the 0x8000 flag', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_SELECTOBJECT, [0x8004 - 0x10000]), // BLACK_BRUSH, as a signed 16-bit value
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#000000');
  });
});

// ---------------------------------------------------------------------------
// Device context stack
// ---------------------------------------------------------------------------

describe('parseWmfVector — device context', () => {
  it('restores the brush saved before a change', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x0000ff)), // red
        record(META_SELECTOBJECT, [0]),
        record(META_SAVEDC),
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x00ff00)), // green
        record(META_SELECTOBJECT, [1]),
        record(META_RESTOREDC, [-1]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#FF0000');
  });
});

// ---------------------------------------------------------------------------
// Robustness
// ---------------------------------------------------------------------------

describe('parseWmfVector — malformed input', () => {
  it('stops at a record that runs past the end of the data', () => {
    const wmf = buildWmf([...windowRecords, rawRecord(META_POLYGON, polyBody([[0, 0], [10, 10]]))]);
    const view = new DataView(wmf.buffer);
    // Claim the polygon record is enormous.
    view.setUint32(22 + 18 + windowRecords[0].length + windowRecords[1].length, 0xffff, true);
    expect(parseWmfVector(wmf)).toBeNull();
  });

  it('stops at a record smaller than the minimum of three words', () => {
    const wmf = buildWmf([...windowRecords, rawRecord(META_POLYGON, polyBody([[0, 0], [10, 10]]))]);
    const view = new DataView(wmf.buffer);
    view.setUint32(22 + 18, 1, true);
    expect(parseWmfVector(wmf)).toBeNull();
  });

  it('ignores a poly record whose point count exceeds its payload', () => {
    const body = new Uint8Array(2 + 4);
    const view = new DataView(body.buffer);
    view.setUint16(0, 500, true); // claims 500 points, carries one
    const image = parseWmfVector(buildWmf([...windowRecords, rawRecord(META_POLYGON, body)]));
    expect(image).toBeNull();
  });

  it('skips unrecognised records and keeps drawing', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_DIBBITBLT, [1, 2, 3, 4]), // bitmap blit — not supported
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Text output
// ---------------------------------------------------------------------------

describe('parseWmfVector text', () => {
  /** Select a font into slot 0 so the records after it draw with it. */
  const withFont = (options?: Parameters<typeof fontBody>[0]) => [
    rawRecord(META_CREATEFONTINDIRECT, fontBody(options)),
    record(META_SELECTOBJECT, [0]),
  ];

  it('draws an ExtTextOut run with the selected font and colour', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont({ height: -20, facename: 'Arial' }),
        record(META_SETTEXTCOLOR, [0x00ff, 0x0000]), // COLORREF 0x000000FF: red
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'))),
      ]),
    );
    expect(image!.texts).toHaveLength(1);
    const [text] = image!.texts!;
    expect(text.text).toBe('Hi');
    expect(text.fill).toBe('#FF0000');
    expect(text.fontSize).toBe(20);
    expect(text.fontFamily).toContain('Arial');
    expect(text.x).toBe(10);
    // Default alignment is TA_TOP, so the baseline sits an ascent lower.
    expect(text.y).toBeCloseTo(37.6, 5);
  });

  it('reads a TextOut string from between the length and the position', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont({ height: -20 }),
        rawRecord(META_TEXTOUT, textOutBody(40, 50, ascii('Hi'))),
      ]),
    );
    expect(image!.texts![0].text).toBe('Hi');
    expect(image!.texts![0].x).toBe(40);
  });

  it('places each glyph at the position its Dx entry gives', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont(),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'), [30, 40])),
      ]),
    );
    expect(image!.texts![0].xs).toEqual([10, 40]);
  });

  it('turns a positive LOGFONT height into an em size', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        // A positive height is the cell, ascent plus descent, not the em.
        ...withFont({ height: 20 }),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'))),
      ]),
    );
    expect(image!.texts![0].fontSize).toBeCloseTo(20, 5);
    expect(image!.texts![0].fontSize).toBeLessThan(20.001);
  });

  it('decodes a double-byte run with the charset the font declares', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont({ charset: 134, facename: 'SimSun' }),
        // GB2312 bytes for 中文, with the per-byte Dx array GDI records.
        rawRecord(
          META_EXTTEXTOUT,
          extTextOutBody(10, 20, Uint8Array.from([0xd6, 0xd0, 0xce, 0xc4]), [30, 0, 40, 0]),
        ),
      ]),
    );
    const [text] = image!.texts!;
    expect(text.text).toBe('中文');
    // Two Dx entries per character have to fold onto one glyph each.
    expect(text.xs).toEqual([10, 40]);
  });

  it('maps Symbol font slots to their Unicode equivalents', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont({ charset: 2, facename: 'Symbol' }),
        // 0xec..0xef are the four pieces a tall curly brace is built from.
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, Uint8Array.from([0xec, 0xef, 0xee]))),
      ]),
    );
    const [text] = image!.texts!;
    expect(text.text).toBe('\u23a7\u23aa\u23a9');
    expect(text.fontFamily).toContain('Cambria Math');
  });

  it('keeps a non-Symbol symbol font in the private-use area it maps', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont({ charset: 2, facename: 'Wingdings' }),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, Uint8Array.from([0x6c]))),
      ]),
    );
    expect(image!.texts![0].text).toBe('\uf06c');
    expect(image!.texts![0].fontFamily).toContain('Wingdings');
  });

  it('takes the reference point from the current position under TA_UPDATECP', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_SETTEXTALIGN, [0x0019]), // TA_UPDATECP | TA_BASELINE
        ...withFont(),
        record(META_MOVETO, [60, 50]), // reverse order: y, x
        rawRecord(META_EXTTEXTOUT, extTextOutBody(0, 0, ascii('A'), [30])),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(0, 0, ascii('B'), [30])),
      ]),
    );
    const [first, second] = image!.texts!;
    expect([first.x, first.y]).toEqual([50, 60]); // TA_BASELINE needs no offset
    // The run advanced the current position, so the next one starts after it.
    expect([second.x, second.y]).toEqual([80, 60]);
  });

  it('folds a centred run into its glyph positions', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_SETTEXTALIGN, [0x0006]), // TA_CENTER
        ...withFont(),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(100, 20, ascii('Hi'), [30, 40])),
      ]),
    );
    const [text] = image!.texts!;
    // Per-glyph positions make every glyph its own text chunk, so `text-anchor`
    // could not centre the run: the 70-unit advance is halved here instead.
    expect(text.xs).toEqual([65, 95]);
    expect(text.anchor).toBe('middle');
  });

  it('leaves alignment to text-anchor when the record carries no Dx array', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        record(META_SETTEXTALIGN, [0x0002]), // TA_RIGHT
        ...withFont(),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(100, 20, ascii('Hi'))),
      ]),
    );
    const [text] = image!.texts!;
    expect(text.xs).toBeUndefined();
    expect(text.anchor).toBe('end');
    expect(text.x).toBe(100);
  });

  it('turns escapement into a clockwise SVG rotation', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        // 900 tenths of a degree counter-clockwise.
        ...withFont({ escapement: 900 }),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'))),
      ]),
    );
    expect(image!.texts![0].rotation).toBe(-90);
  });

  it('skips the rectangle an opaque or clipped run prefixes its string with', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont(),
        // ETO_CLIPPED, so the string is preceded by a clipping rectangle.
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'), undefined, 0x0004)),
      ]),
    );
    expect(image!.texts![0].text).toBe('Hi');
  });

  it('selects a font without disturbing the brush in the same table', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(META_CREATEBRUSHINDIRECT, brushBody(0, 0x00ff00)), // slot 0
        rawRecord(META_CREATEFONTINDIRECT, fontBody()), // slot 1
        record(META_SELECTOBJECT, [0]),
        record(META_SELECTOBJECT, [1]),
        record(META_RECTANGLE, [50, 50, 0, 0]),
      ]),
    );
    expect(image!.paths[0].fill).toBe('#00FF00');
  });

  it('frames a file whose only content is text', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        ...withFont(),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('Hi'))),
      ]),
    );
    // The bug this guards: a text-only metafile used to convert to nothing, so
    // the picture was a hole in the slide.
    expect(image).not.toBeNull();
    expect(image!.paths).toHaveLength(0);
    expect([image!.width, image!.height]).toEqual([200, 100]);
  });
});

// ---------------------------------------------------------------------------
// SVG serialization
// ---------------------------------------------------------------------------

describe('wmfVectorToSvg', () => {
  it('emits a sized, self-contained SVG document', () => {
    const image = parseWmfVector(
      buildWmf([...windowRecords, record(META_RECTANGLE, [40, 60, 10, 20])]),
    );
    const svg = wmfVectorToSvg(image!);
    expect(svg).toContain('<svg xmlns="http://www.w3.org/2000/svg"');
    expect(svg).toContain('width="200" height="100"');
    expect(svg).toContain('viewBox="0 0 200 100"');
    expect(svg).toContain('<path d="M20,10 L60,10 L60,40 L20,40 Z"');
  });

  it('emits a <text> element carrying the run\'s font and decorations', () => {
    const image = parseWmfVector(
      buildWmf([
        ...windowRecords,
        rawRecord(
          META_CREATEFONTINDIRECT,
          fontBody({ height: -20, weight: 700, italic: true, underline: true, facename: 'Arial' }),
        ),
        record(META_SELECTOBJECT, [0]),
        rawRecord(META_EXTTEXTOUT, extTextOutBody(10, 20, ascii('a & b'), [4, 4, 4, 4, 4])),
      ]),
    );
    const svg = wmfVectorToSvg(image!);
    expect(svg).toContain('<text x="10 14 18 22 26" y="37.6"');
    expect(svg).toContain('font-size="20"');
    expect(svg).toContain('font-weight="700"');
    expect(svg).toContain('font-style="italic"');
    expect(svg).toContain('text-decoration="underline"');
    expect(svg).toContain('>a &amp; b</text>');
  });
});
