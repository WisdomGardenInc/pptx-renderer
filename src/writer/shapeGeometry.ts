import type { CustomShapeGeometry, ShapeGeometryCommand } from '../shapes/editingGeometry';
import {
  A,
  P,
  SP_PR_ORDER,
  XFRM_ORDER,
  childByLocal,
  ensureChild,
  ensureShapeProperties,
  insertOrdered,
} from './xmlEdit';

export type ShapeGeometry =
  | { readonly kind: 'preset'; readonly preset: string }
  | ({ readonly kind: 'custom' } & CustomShapeGeometry);
export interface ShapeGeometryOptions {
  readonly setExtent?: boolean;
  readonly convertConnector?: boolean;
}
export interface ShapeElementOptions {
  readonly id: string;
  readonly name?: string;
  readonly preset?: string;
  readonly connector?: boolean;
}

function append(
  parent: Element,
  name: string,
  attributes: Record<string, string | number> = {},
): Element {
  const element = parent.ownerDocument.createElementNS(A, `a:${name}`);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, String(value));
  parent.appendChild(element);
  return element;
}

function validateCommand(command: ShapeGeometryCommand): void {
  if (command.kind === 'arc') {
    const { radiusX, radiusY, startAngle, sweepAngle } = command;
    if (
      ![radiusX, radiusY, startAngle, sweepAngle].every(Number.isFinite) ||
      radiusX < 0 ||
      radiusY < 0
    )
      throw new RangeError('Arc dimensions and angles must be finite.');
    return;
  }
  const lengths = { M: 2, L: 2, Q: 4, C: 6, Z: 0 };
  if (command.points.length !== lengths[command.kind] || !command.points.every(Number.isFinite))
    throw new RangeError('Invalid geometry path coordinates.');
}

function validateGeometry(geometry: ShapeGeometry): void {
  if (geometry.kind === 'preset') {
    if (!geometry.preset) throw new RangeError('Preset geometry requires a name.');
    return;
  }
  if (![geometry.width, geometry.height].every((value) => Number.isFinite(value) && value >= 0))
    throw new RangeError('Geometry extents must be finite and nonnegative.');
  for (const path of geometry.paths) for (const command of path.commands) validateCommand(command);
}

function writeCommand(path: Element, command: ShapeGeometryCommand): void {
  if (command.kind === 'arc') {
    append(path, 'arcTo', {
      wR: Math.round(command.radiusX * 9525),
      hR: Math.round(command.radiusY * 9525),
      stAng: Math.round(command.startAngle * 60000),
      swAng: Math.round(command.sweepAngle * 60000),
    });
    return;
  }
  const names = { M: 'moveTo', L: 'lnTo', Q: 'quadBezTo', C: 'cubicBezTo', Z: 'close' };
  const node = append(path, names[command.kind]);
  for (let index = 0; index < command.points.length; index += 2)
    append(node, 'pt', {
      x: Math.round(command.points[index] * 9525),
      y: Math.round(command.points[index + 1] * 9525),
    });
}

function customGeometry(properties: Element, geometry: CustomShapeGeometry): void {
  const node = properties.ownerDocument.createElementNS(A, 'a:custGeom');
  insertOrdered(properties, node, SP_PR_ORDER);
  for (const name of ['avLst', 'gdLst', 'ahLst', 'cxnLst']) append(node, name);
  append(node, 'rect', { l: '0', t: '0', r: 'r', b: 'b' });
  const paths = append(node, 'pathLst');
  for (const path of geometry.paths) {
    const element = append(paths, 'path', {
      w: Math.max(1, Math.round(geometry.width * 9525)),
      h: Math.max(1, Math.round(geometry.height * 9525)),
    });
    if (path.fill !== undefined) element.setAttribute('fill', path.fill);
    if (path.stroke !== undefined) element.setAttribute('stroke', String(path.stroke));
    for (const command of path.commands) writeCommand(element, command);
  }
}

function convertConnector(source: Element): Element {
  if (source.localName !== 'cxnSp') return source;
  const document = source.ownerDocument;
  const shape = document.createElementNS(P, 'p:sp');
  for (const attribute of Array.from(source.attributes))
    shape.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
  const old = childByLocal(source, 'nvCxnSpPr');
  if (old) {
    const nv = document.createElementNS(P, 'p:nvSpPr');
    for (const node of Array.from(old.children)) {
      if (node.localName !== 'cNvCxnSpPr') {
        nv.appendChild(node);
        continue;
      }
      const properties = document.createElementNS(P, 'p:cNvSpPr');
      for (const attribute of Array.from(node.attributes))
        properties.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      const locks = childByLocal(node, 'cxnSpLocks');
      if (locks) {
        const shapeLocks = append(properties, 'spLocks');
        for (const attribute of Array.from(locks.attributes).filter(
          (value) => value.localName !== 'noDisconnect',
        ))
          shapeLocks.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      }
      const extensions = childByLocal(node, 'extLst');
      if (extensions) properties.appendChild(extensions);
      nv.appendChild(properties);
    }
    old.replaceWith(nv);
  }
  shape.append(...Array.from(source.childNodes));
  if (source.parentNode) source.replaceWith(shape);
  return shape;
}

export function writeShapeGeometry(
  source: Element,
  geometry: ShapeGeometry,
  options: ShapeGeometryOptions = {},
): Element {
  validateGeometry(geometry);
  const properties = ensureShapeProperties(source);
  childByLocal(properties, 'prstGeom')?.remove();
  childByLocal(properties, 'custGeom')?.remove();
  if (geometry.kind === 'preset') {
    const preset = ensureChild(properties, A, 'a:prstGeom', SP_PR_ORDER);
    preset.setAttribute('prst', geometry.preset);
    append(preset, 'avLst');
  } else {
    customGeometry(properties, geometry);
    if (options.setExtent) {
      const transform = ensureChild(properties, A, 'a:xfrm', SP_PR_ORDER);
      const ext = ensureChild(transform, A, 'a:ext', XFRM_ORDER);
      ext.setAttribute('cx', String(Math.round(geometry.width * 9525)));
      ext.setAttribute('cy', String(Math.round(geometry.height * 9525)));
    }
  }
  return options.convertConnector ? convertConnector(source) : source;
}

export function createShapeElement(document: Document, options: ShapeElementOptions): Element {
  const tag = options.connector ? 'cxnSp' : 'sp';
  const source = document.createElementNS(P, `p:${tag}`);
  const nv = ensureChild(source, P, options.connector ? 'p:nvCxnSpPr' : 'p:nvSpPr', []);
  const properties = ensureChild(nv, P, 'p:cNvPr', []);
  properties.setAttribute('id', options.id);
  properties.setAttribute('name', options.name ?? 'Shape');
  ensureChild(nv, P, options.connector ? 'p:cNvCxnSpPr' : 'p:cNvSpPr', []);
  ensureChild(nv, P, 'p:nvPr', []);
  const shapeProperties = ensureShapeProperties(source);
  const transform = ensureChild(shapeProperties, A, 'a:xfrm', SP_PR_ORDER);
  const off = ensureChild(transform, A, 'a:off', XFRM_ORDER);
  off.setAttribute('x', '0');
  off.setAttribute('y', '0');
  const ext = ensureChild(transform, A, 'a:ext', XFRM_ORDER);
  ext.setAttribute('cx', '0');
  ext.setAttribute('cy', '0');
  return writeShapeGeometry(source, { kind: 'preset', preset: options.preset ?? 'rect' });
}
