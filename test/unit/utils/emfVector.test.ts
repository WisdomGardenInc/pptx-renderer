import { describe, it, expect } from 'vitest';
import { parseEmfVector, emfVectorToSvg } from '../../../src/utils/emfVector';

// ---------------------------------------------------------------------------
// Helpers to build synthetic EMF binary data
// ---------------------------------------------------------------------------

const EMF_SIGNATURE = 0x464d4520;

/** Build a record: 8-byte header plus a body of little-endian int32 words. */
function record(type: number, words: number[] = [], extra?: Uint8Array): Uint8Array {
  const extraBytes = extra ? extra.length : 0;
  const size = 8 + words.length * 4 + extraBytes;
  const buf = new Uint8Array(size);
  const view = new DataView(buf.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, size, true);
  words.forEach((w, i) => view.setInt32(8 + i * 4, w | 0, true));
  if (extra) buf.set(extra, 8 + words.length * 4);
  return buf;
}

/** Pack 16-bit points as the poly*16 records store them. */
function points16(pts: Array<[number, number]>): Uint8Array {
  const buf = new Uint8Array(pts.length * 4);
  const view = new DataView(buf.buffer);
  pts.forEach(([x, y], i) => {
    view.setInt16(i * 4, x, true);
    view.setInt16(i * 4 + 2, y, true);
  });
  return buf;
}

/** Pack 32-bit points as the poly* records store them. */
function points32(pts: Array<[number, number]>): Uint8Array {
  const buf = new Uint8Array(pts.length * 8);
  const view = new DataView(buf.buffer);
  pts.forEach(([x, y], i) => {
    view.setInt32(i * 8, x, true);
    view.setInt32(i * 8 + 4, y, true);
  });
  return buf;
}

interface HeaderOptions {
  bounds?: [number, number, number, number];
  frame?: [number, number, number, number];
  device?: [number, number];
  millimeters?: [number, number];
}

/**
 * Build an 88-byte EMF header record.
 *
 * Defaults describe a 200x100 device-unit drawing on a 1024x768 / 320x240 mm
 * surface (3.2 px per mm), so the frame in 0.01 mm maps exactly onto the bounds.
 */
function buildHeader(options: HeaderOptions = {}): Uint8Array {
  const {
    bounds = [0, 0, 200, 100],
    frame = [0, 0, 6250, 3125],
    device = [1024, 768],
    millimeters = [320, 240],
  } = options;
  const buf = new Uint8Array(88);
  const view = new DataView(buf.buffer);
  view.setUint32(0, 1, true); // EMR_HEADER
  view.setUint32(4, 88, true);
  bounds.forEach((v, i) => view.setInt32(8 + i * 4, v, true));
  frame.forEach((v, i) => view.setInt32(24 + i * 4, v, true));
  view.setUint32(40, EMF_SIGNATURE, true);
  view.setInt32(72, device[0], true);
  view.setInt32(76, device[1], true);
  view.setInt32(80, millimeters[0], true);
  view.setInt32(84, millimeters[1], true);
  return buf;
}

const eofRecord = () => record(14, [0, 0, 20]);

function buildEmf(header: Uint8Array, records: Uint8Array[]): Uint8Array {
  const parts = [header, ...records, eofRecord()];
  const total = parts.reduce((sum, p) => sum + p.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

// Record builders used across the tests.
const setMapMode = (mode: number) => record(17, [mode]);
const setWindowOrg = (x: number, y: number) => record(10, [x, y]);
const setWindowExt = (x: number, y: number) => record(9, [x, y]);
const setViewportOrg = (x: number, y: number) => record(12, [x, y]);
const setViewportExt = (x: number, y: number) => record(11, [x, y]);
const setPolyFillMode = (mode: number) => record(19, [mode]);
const createBrush = (handle: number, style: number, colorRef: number) =>
  record(39, [handle, style, colorRef, 0]);
const createPen = (handle: number, style: number, width: number, colorRef: number) =>
  record(38, [handle, style, width, 0, colorRef]);
const extCreatePen = (handle: number, style: number, width: number, colorRef: number) =>
  record(95, [handle, 44, 0, 44, 0, style, width, 0, colorRef, 0, 0]);
const selectObject = (handle: number) => record(37, [handle]);
const beginPath = () => record(59);
const endPath = () => record(60);
const closeFigure = () => record(61);
const fillPath = () => record(62, [0, 0, 0, 0]);
const strokePath = () => record(64, [0, 0, 0, 0]);
const strokeAndFillPath = () => record(63, [0, 0, 0, 0]);
const moveTo = (x: number, y: number) => record(27, [x, y]);
const lineTo = (x: number, y: number) => record(54, [x, y]);
const polylineTo16 = (pts: Array<[number, number]>) =>
  record(89, [0, 0, 0, 0, pts.length], points16(pts));
const polyBezierTo16 = (pts: Array<[number, number]>) =>
  record(88, [0, 0, 0, 0, pts.length], points16(pts));
const polygon16 = (pts: Array<[number, number]>) =>
  record(86, [0, 0, 0, 0, pts.length], points16(pts));
const polygon32 = (pts: Array<[number, number]>) =>
  record(3, [0, 0, 0, 0, pts.length], points32(pts));
const polyPolygon16 = (polys: Array<Array<[number, number]>>) => {
  const flat = polys.flat();
  return record(
    91,
    [0, 0, 0, 0, polys.length, flat.length, ...polys.map((p) => p.length)],
    points16(flat),
  );
};
const rectangle = (l: number, t: number, r: number, b: number) => record(43, [l, t, r, b]);
const ellipse = (l: number, t: number, r: number, b: number) => record(42, [l, t, r, b]);
const saveDc = () => record(33);
const restoreDc = (level: number) => record(34, [level]);
const setWorldTransform = (m: [number, number, number, number, number, number]) => {
  const body = new Uint8Array(24);
  const view = new DataView(body.buffer);
  m.forEach((v, i) => view.setFloat32(i * 4, v, true));
  return record(35, [], body);
};

const MM_TEXT = 1;
const MM_ANISOTROPIC = 8;
const BS_SOLID = 0;
const BS_NULL = 1;
const PS_SOLID = 0;
const PS_NULL = 5;
const ALTERNATE = 1;
const WINDING = 2;

/** A minimal filled triangle drawn through the path bracket. */
function filledTriangle(colorRef: number, fillMode = WINDING): Uint8Array[] {
  return [
    createBrush(1, BS_SOLID, colorRef),
    selectObject(1),
    setPolyFillMode(fillMode),
    beginPath(),
    moveTo(0, 0),
    polylineTo16([
      [10, 0],
      [10, 10],
    ]),
    closeFigure(),
    endPath(),
    fillPath(),
  ];
}

// ---------------------------------------------------------------------------

describe('parseEmfVector', () => {
  describe('validation', () => {
    it('returns null for data too small to hold a header', () => {
      expect(parseEmfVector(new Uint8Array(40))).toBeNull();
    });

    it('returns null when the EMF signature is missing', () => {
      const data = buildEmf(buildHeader(), filledTriangle(0x000000));
      new DataView(data.buffer).setUint32(40, 0xdeadbeef, true);
      expect(parseEmfVector(data)).toBeNull();
    });

    it('returns null for a header + EOF file with no geometry', () => {
      expect(parseEmfVector(buildEmf(buildHeader(), []))).toBeNull();
    });

    it('returns null when records draw nothing paintable', () => {
      const data = buildEmf(buildHeader(), [setMapMode(MM_TEXT), setWindowOrg(0, 0)]);
      expect(parseEmfVector(data)).toBeNull();
    });
  });

  describe('viewBox', () => {
    it('derives the viewBox from rclFrame scaled by the device metrics', () => {
      // 3.2 px/mm; frame 0..6250 (0.01mm) → 0..62.5mm → 0..200 device units.
      const image = parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x000000)));
      expect(image).not.toBeNull();
      expect(image!.x).toBeCloseTo(0, 6);
      expect(image!.y).toBeCloseTo(0, 6);
      expect(image!.width).toBeCloseTo(200, 6);
      expect(image!.height).toBeCloseTo(100, 6);
    });

    it('keeps a non-zero frame origin', () => {
      const header = buildHeader({ frame: [3125, 1562.5 | 0, 6250, 3125] });
      const image = parseEmfVector(buildEmf(header, filledTriangle(0x000000)));
      expect(image!.x).toBeCloseTo(100, 6);
      expect(image!.width).toBeCloseTo(100, 6);
    });

    it('falls back to rclBounds when the frame is degenerate', () => {
      const header = buildHeader({ bounds: [5, 7, 45, 27], frame: [0, 0, 0, 0] });
      const image = parseEmfVector(buildEmf(header, filledTriangle(0x000000)));
      expect(image).toMatchObject({ x: 5, y: 7, width: 40, height: 20 });
    });

    it('falls back to rclBounds when device metrics are missing', () => {
      const header = buildHeader({ bounds: [0, 0, 40, 20], millimeters: [0, 0] });
      const image = parseEmfVector(buildEmf(header, filledTriangle(0x000000)));
      expect(image).toMatchObject({ x: 0, y: 0, width: 40, height: 20 });
    });

    it('returns null when neither frame nor bounds describe an area', () => {
      const header = buildHeader({ bounds: [0, 0, 0, 0], frame: [0, 0, 0, 0] });
      expect(parseEmfVector(buildEmf(header, filledTriangle(0x000000)))).toBeNull();
    });
  });

  describe('brushes and pens', () => {
    it('fills with the brush color, decoding COLORREF as 0x00bbggrr', () => {
      // 0x0000FF is pure red in COLORREF byte order.
      const image = parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x0000ff)));
      expect(image!.paths).toHaveLength(1);
      expect(image!.paths[0].fill).toBe('#FF0000');
      expect(image!.paths[0].stroke).toBe('none');
    });

    it('maps ALTERNATE to evenodd and WINDING to nonzero', () => {
      const alt = parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x000000, ALTERNATE)));
      const wind = parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x000000, WINDING)));
      expect(alt!.paths[0].fillRule).toBe('evenodd');
      expect(wind!.paths[0].fillRule).toBe('nonzero');
    });

    it('drops the fill for a BS_NULL (hollow) brush', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_NULL, 0x0000ff),
        selectObject(1),
        createPen(2, PS_SOLID, 3, 0x00ff00),
        selectObject(2),
        polygon16([
          [0, 0],
          [10, 0],
          [10, 10],
        ]),
      ]);
      const image = parseEmfVector(data);
      expect(image!.paths[0].fill).toBe('none');
      expect(image!.paths[0].stroke).toBe('#00FF00');
    });

    it('drops the stroke for a PS_NULL pen created by EXTCREATEPEN', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x0000ff),
        selectObject(1),
        extCreatePen(2, PS_NULL, 5, 0x00ff00),
        selectObject(2),
        polygon16([
          [0, 0],
          [10, 0],
          [10, 10],
        ]),
      ]);
      const image = parseEmfVector(data);
      expect(image!.paths[0].fill).toBe('#FF0000');
      expect(image!.paths[0].stroke).toBe('none');
    });

    it('scales the pen width into device units', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_ANISOTROPIC),
        setWindowExt(100, 100),
        setViewportExt(50, 50), // 0.5x
        createPen(1, PS_SOLID, 8, 0x000000),
        selectObject(1),
        beginPath(),
        moveTo(0, 0),
        lineTo(10, 10),
        endPath(),
        strokePath(),
      ]);
      const image = parseEmfVector(data);
      expect(image!.paths[0].strokeWidth).toBeCloseTo(4, 6);
    });

    it('honours the NULL_BRUSH and BLACK_PEN stock objects', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x0000ff),
        selectObject(1),
        selectObject(0x80000005), // NULL_BRUSH
        selectObject(0x80000007), // BLACK_PEN
        polygon16([
          [0, 0],
          [10, 0],
          [10, 10],
        ]),
      ]);
      const image = parseEmfVector(data);
      expect(image!.paths[0].fill).toBe('none');
      expect(image!.paths[0].stroke).toBe('#000000');
    });

    it('emits nothing when both brush and pen are invisible', () => {
      const data = buildEmf(buildHeader(), [
        selectObject(0x80000005), // NULL_BRUSH
        selectObject(0x80000008), // NULL_PEN
        polygon16([
          [0, 0],
          [10, 0],
          [10, 10],
        ]),
      ]);
      expect(parseEmfVector(data)).toBeNull();
    });
  });

  describe('coordinate mapping', () => {
    it('treats MM_TEXT as an identity mapping and ignores window/viewport extents', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_TEXT),
        setWindowExt(100, 100),
        setViewportExt(50, 50),
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M0,0 L10,0 L10,10 Z');
    });

    it('applies the window/viewport ratio under MM_ANISOTROPIC', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_ANISOTROPIC),
        setWindowOrg(0, 0),
        setWindowExt(100, 100),
        setViewportOrg(0, 0),
        setViewportExt(50, 200), // 0.5x horizontally, 2x vertically
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M0,0 L5,0 L5,20 Z');
    });

    it('offsets by the window and viewport origins', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_ANISOTROPIC),
        setWindowOrg(10, 10),
        setWindowExt(1, 1),
        setViewportOrg(100, 200),
        setViewportExt(1, 1),
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M90,190 L100,190 L100,200 Z');
    });

    it('applies a world transform before the map-mode mapping', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_TEXT),
        setWorldTransform([2, 0, 0, 2, 5, 5]), // scale 2x then translate
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M5,5 L25,5 L25,25 Z');
    });

    it('restores the transform state saved by SAVEDC', () => {
      const data = buildEmf(buildHeader(), [
        setMapMode(MM_TEXT),
        saveDc(),
        setWorldTransform([2, 0, 0, 2, 0, 0]),
        restoreDc(-1),
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M0,0 L10,0 L10,10 Z');
    });
  });

  describe('geometry records', () => {
    it('builds cubic segments from POLYBEZIERTO16 continuing the current point', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        beginPath(),
        moveTo(0, 0),
        polyBezierTo16([
          [1, 2],
          [3, 4],
          [5, 6],
        ]),
        closeFigure(),
        endPath(),
        fillPath(),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M0,0 C1,2 3,4 5,6 Z');
    });

    it('reads 32-bit point records as well as the 16-bit variants', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        polygon32([
          [0, 0],
          [40000, 0],
          [40000, 40000],
        ]),
      ]);
      // 40000 overflows int16, so this only works with the 32-bit reader.
      expect(parseEmfVector(data)!.paths[0].d).toBe('M0,0 L40000,0 L40000,40000 Z');
    });

    it('emits one path per POLYPOLYGON16, with a subpath per polygon', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        polyPolygon16([
          [
            [0, 0],
            [10, 0],
            [10, 10],
          ],
          [
            [20, 20],
            [30, 20],
            [30, 30],
          ],
        ]),
      ]);
      const paths = parseEmfVector(data)!.paths;
      expect(paths).toHaveLength(1);
      expect(paths[0].d).toBe('M0,0 L10,0 L10,10 Z M20,20 L30,20 L30,30 Z');
    });

    it('renders RECTANGLE as a closed quad', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        rectangle(2, 4, 12, 14),
      ]);
      expect(parseEmfVector(data)!.paths[0].d).toBe('M2,4 L12,4 L12,14 L2,14 Z');
    });

    it('renders ELLIPSE as four cubic arcs', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        ellipse(0, 0, 20, 20),
      ]);
      const d = parseEmfVector(data)!.paths[0].d;
      expect(d.startsWith('M0,10 ')).toBe(true);
      expect(d.match(/C/g)).toHaveLength(4);
      expect(d.endsWith('Z')).toBe(true);
    });

    it('strokes a LINETO issued outside a path bracket', () => {
      const data = buildEmf(buildHeader(), [
        createPen(1, PS_SOLID, 1, 0x000000),
        selectObject(1),
        moveTo(1, 1),
        lineTo(5, 5),
      ]);
      const paths = parseEmfVector(data)!.paths;
      expect(paths).toHaveLength(1);
      expect(paths[0].d).toBe('M1,1 L5,5');
      expect(paths[0].fill).toBe('none');
    });

    it('accumulates a whole path bracket into a single painted path', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        beginPath(),
        moveTo(0, 0),
        lineTo(10, 0),
        closeFigure(),
        moveTo(20, 20),
        lineTo(30, 20),
        closeFigure(),
        endPath(),
        fillPath(),
      ]);
      const paths = parseEmfVector(data)!.paths;
      expect(paths).toHaveLength(1);
      expect(paths[0].d).toBe('M0,0 L10,0 Z M20,20 L30,20 Z');
    });

    it('paints both fill and stroke for STROKEANDFILLPATH', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x0000ff),
        selectObject(1),
        createPen(2, PS_SOLID, 2, 0x00ff00),
        selectObject(2),
        beginPath(),
        moveTo(0, 0),
        lineTo(10, 10),
        closeFigure(),
        endPath(),
        strokeAndFillPath(),
      ]);
      const path = parseEmfVector(data)!.paths[0];
      expect(path.fill).toBe('#FF0000');
      expect(path.stroke).toBe('#00FF00');
    });

    it('discards the pending path on ABORTPATH', () => {
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        beginPath(),
        moveTo(0, 0),
        lineTo(10, 10),
        record(68), // EMR_ABORTPATH
        fillPath(),
        ...filledTriangle(0x000000),
      ]);
      const paths = parseEmfVector(data)!.paths;
      expect(paths).toHaveLength(1);
      expect(paths[0].d).toBe('M0,0 L10,0 L10,10 Z');
    });
  });

  describe('robustness', () => {
    it('keeps the geometry parsed before a corrupted record size', () => {
      const good = buildEmf(buildHeader(), filledTriangle(0x000000));
      const broken = new Uint8Array(good.length + 8);
      broken.set(good.subarray(0, good.length - 20)); // drop the EOF record
      const view = new DataView(broken.buffer);
      view.setUint32(good.length - 20, 86, true); // POLYGON16
      view.setUint32(good.length - 16, 0xfffffff0, true); // absurd size
      expect(parseEmfVector(broken)!.paths).toHaveLength(1);
    });

    it('ignores unsupported records such as text output', () => {
      const data = buildEmf(buildHeader(), [
        record(84, [0, 0, 0, 0, 0, 0, 0, 0]), // EMR_EXTTEXTOUTW
        ...filledTriangle(0x000000),
      ]);
      expect(parseEmfVector(data)!.paths).toHaveLength(1);
    });

    it('rejects poly records whose point count exceeds the payload', () => {
      const bogus = record(86, [0, 0, 0, 0, 1000], points16([[0, 0]]));
      const data = buildEmf(buildHeader(), [
        createBrush(1, BS_SOLID, 0x000000),
        selectObject(1),
        bogus,
      ]);
      expect(parseEmfVector(data)).toBeNull();
    });

    it('handles a Uint8Array that is a subarray of a larger buffer', () => {
      const source = buildEmf(buildHeader(), filledTriangle(0x0000ff));
      const padded = new Uint8Array(source.length + 16);
      padded.set(source, 8);
      const image = parseEmfVector(padded.subarray(8, 8 + source.length));
      expect(image!.paths[0].fill).toBe('#FF0000');
    });
  });
});

describe('emfVectorToSvg', () => {
  it('serializes a drawing with an intrinsic size and a matching viewBox', () => {
    const svg = emfVectorToSvg(parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x0000ff)))!);
    expect(svg).toContain('xmlns="http://www.w3.org/2000/svg"');
    expect(svg).toContain('width="200"');
    expect(svg).toContain('height="100"');
    expect(svg).toContain('viewBox="0 0 200 100"');
    expect(svg).toContain('preserveAspectRatio="none"');
    expect(svg).toContain('<path d="M0,0 L10,0 L10,10 Z" fill="#FF0000"/>');
  });

  it('emits fill-rule only for evenodd fills', () => {
    const evenOdd = emfVectorToSvg(
      parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x000000, ALTERNATE)))!,
    );
    const nonZero = emfVectorToSvg(
      parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x000000, WINDING)))!,
    );
    expect(evenOdd).toContain('fill-rule="evenodd"');
    expect(nonZero).not.toContain('fill-rule');
  });

  it('renders a zero-width GDI pen as a one-unit cosmetic stroke', () => {
    const data = buildEmf(buildHeader(), [
      createPen(1, PS_SOLID, 0, 0x000000),
      selectObject(1),
      moveTo(0, 0),
      lineTo(10, 10),
    ]);
    expect(emfVectorToSvg(parseEmfVector(data)!)).toContain('stroke-width="1"');
  });

  it('produces valid XML that a DOM parser accepts', () => {
    const svg = emfVectorToSvg(parseEmfVector(buildEmf(buildHeader(), filledTriangle(0x0000ff)))!);
    const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
    expect(doc.querySelector('parsererror')).toBeNull();
    expect(doc.documentElement.tagName).toBe('svg');
    expect(doc.querySelectorAll('path')).toHaveLength(1);
  });
});
