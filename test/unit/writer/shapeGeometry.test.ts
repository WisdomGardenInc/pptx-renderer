import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseShapeNode } from '../../../src/model/nodes/ShapeNode';
import { createShapeElement, writeShapeGeometry } from '../../../src/writer/shapeGeometry';
import { writeShapePaint } from '../../../src/writer/shapePaint';
import {
  createFilledLineGeometry,
  createLineEndpointGeometry,
} from '../../../src/shapes/editingGeometry';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const connectorXml = `<p:cxnSp xmlns:p="${P}" xmlns:a="${A}"><p:nvCxnSpPr><p:cNvPr id="7" name="Line"/><p:cNvCxnSpPr><a:cxnSpLocks noSelect="1" noDisconnect="1"/><a:extLst><a:ext uri="retain-locks"/></a:extLst></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm flipH="1"><a:off x="9525" y="19050"/><a:ext cx="2857500" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:noFill/><a:ln w="76200"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:headEnd type="triangle"/></a:ln><a:effectLst><a:glow rad="123"/></a:effectLst><a:extLst><a:ext uri="retain-style"/></a:extLst></p:spPr></p:cxnSp>`;

function connector(): Element {
  return new DOMParser().parseFromString(connectorXml, 'application/xml').documentElement;
}

const none = { kind: 'none' } as const;

describe('custom shape authoring', () => {
  it('converts connectors explicitly while preserving identity, transforms, locks and extensions', () => {
    const original = connector();
    const geometry = createFilledLineGeometry({
      width: 300,
      height: 0,
      strokeWidth: 8,
      start: none,
      end: createLineEndpointGeometry('diamond', { length: 24, width: 24, outlineWidth: 4 }),
      startInset: 0,
      endInset: 0,
    });
    const source = writeShapeGeometry(
      original,
      { kind: 'custom', ...geometry },
      { setExtent: true, convertConnector: true },
    );
    writeShapePaint(source, {
      fill: { color: '#123456', alpha: 0.6 },
      stroke: { color: null, width: 0, cap: 'flat', headEnd: null, tailEnd: null },
    });
    const xml = new XMLSerializer().serializeToString(source);
    const node = parseXml(xml);
    expect(source.localName).toBe('sp');
    expect(source.ownerDocument.documentElement).toBe(source);
    expect(node.child('nvSpPr').child('cNvPr').attr('id')).toBe('7');
    expect(node.child('nvSpPr').child('cNvSpPr').child('spLocks').attr('noSelect')).toBe('1');
    expect(
      node.child('nvSpPr').child('cNvSpPr').child('spLocks').attr('noDisconnect'),
    ).toBeUndefined();
    expect(xml).toContain('retain-locks');
    expect(xml).toContain('retain-style');
    const properties = node.child('spPr');
    expect(properties.child('xfrm').attr('flipH')).toBe('1');
    expect(properties.child('xfrm').child('off').attr('x')).toBe('9525');
    expect(properties.child('xfrm').child('ext').attr('cy')).toBe('9525');
    expect(properties.child('effectLst').child('glow').attr('rad')).toBe('123');
    expect(properties.child('ln').child('headEnd').exists()).toBe(false);
    expect(properties.child('solidFill').child('srgbClr').child('alpha').attr('val')).toBe('60000');
    expect(Array.from(source.getElementsByTagNameNS(A, 'path'))).toHaveLength(2);
    expect(Array.from(properties.element!.children).map((child) => child.localName)).toEqual([
      'xfrm',
      'custGeom',
      'solidFill',
      'ln',
      'effectLst',
      'extLst',
    ]);
    const parsed = parseShapeNode(node);
    expect(parsed.id).toBe('7');
    expect(parsed.customGeometry?.exists()).toBe(true);
  });

  it('writes curves and arcs in pixels and degrees with schema-ordered paint', () => {
    const document = new DOMParser().parseFromString(
      `<p:sld xmlns:p="${P}" xmlns:a="${A}"/>`,
      'application/xml',
    );
    const source = createShapeElement(document, { id: '3', name: 'Curve', preset: 'rect' });
    document.documentElement.appendChild(source);
    writeShapeGeometry(source, {
      kind: 'custom',
      width: 40,
      height: 20,
      paths: [
        {
          fill: 'none',
          stroke: true,
          commands: [
            { kind: 'M', points: [0, 0] },
            { kind: 'Q', points: [10, 20, 20, 0] },
            { kind: 'C', points: [20, 0, 30, 20, 40, 0] },
            { kind: 'arc', radiusX: 10, radiusY: 5, startAngle: 0, sweepAngle: 180 },
            { kind: 'Z', points: [] },
          ],
        },
      ],
    });
    writeShapePaint(source, {
      fill: { color: null },
      stroke: { color: '#ff0000', width: 2, cap: 'rnd', join: 'round', dash: 'solid' },
    });
    const node = parseXml(new XMLSerializer().serializeToString(source));
    const path = node.child('spPr').child('custGeom').child('pathLst').child('path');
    expect(path.attr('w')).toBe('381000');
    expect(path.child('quadBezTo').children('pt')).toHaveLength(2);
    expect(path.child('cubicBezTo').children('pt')).toHaveLength(3);
    expect(path.child('arcTo').attr('wR')).toBe('95250');
    expect(path.child('arcTo').attr('hR')).toBe('47625');
    expect(path.child('arcTo').attr('swAng')).toBe('10800000');
    expect(node.child('spPr').child('ln').child('round').exists()).toBe(true);
    expect(node.child('spPr').child('ln').attr('w')).toBe('19050');
  });

  it('does not convert connectors or overwrite extents unless requested', () => {
    const source = connector();
    const before = new XMLSerializer().serializeToString(
      source.getElementsByTagNameNS(A, 'xfrm')[0],
    );
    expect(writeShapeGeometry(source, { kind: 'custom', width: 300, height: 1, paths: [] })).toBe(
      source,
    );
    expect(source.localName).toBe('cxnSp');
    expect(new XMLSerializer().serializeToString(source.getElementsByTagNameNS(A, 'xfrm')[0])).toBe(
      before,
    );
    writeShapeGeometry(source, { kind: 'preset', preset: 'triangle' });
    expect(source.getElementsByTagNameNS(A, 'custGeom')).toHaveLength(0);
    expect(source.getElementsByTagNameNS(A, 'prstGeom')[0].getAttribute('prst')).toBe('triangle');
  });

  it('validates shape roots and paint values without modifying unsupported elements', () => {
    const document = new DOMParser().parseFromString(`<p:sld xmlns:p="${P}"/>`, 'application/xml');
    const source = document.documentElement;
    expect(() => writeShapeGeometry(source, { kind: 'preset', preset: 'rect' })).toThrow(TypeError);
    expect(() => writeShapePaint(source, { fill: { color: '#123456' } })).toThrow(TypeError);
    expect(source.children).toHaveLength(0);
    const shape = connector();
    const before = new XMLSerializer().serializeToString(shape);
    expect(() => writeShapePaint(shape, { stroke: { width: -1 } })).toThrow(RangeError);
    expect(() => writeShapePaint(shape, { fill: { color: '#123456', alpha: NaN } })).toThrow(
      RangeError,
    );
    expect(new XMLSerializer().serializeToString(shape)).toBe(before);
  });

  it('rejects non-finite coordinates and invalid command arity without modifying the source', () => {
    const source = connector();
    const before = new XMLSerializer().serializeToString(source);
    expect(() =>
      writeShapeGeometry(source, {
        kind: 'custom',
        width: 10,
        height: 10,
        paths: [{ commands: [{ kind: 'M', points: [NaN, 0] }] }],
      }),
    ).toThrow(RangeError);
    expect(() =>
      writeShapeGeometry(source, {
        kind: 'custom',
        width: 10,
        height: 10,
        paths: [{ commands: [{ kind: 'Q', points: [1, 2] }] }],
      }),
    ).toThrow(RangeError);
    expect(new XMLSerializer().serializeToString(source)).toBe(before);
  });
});
