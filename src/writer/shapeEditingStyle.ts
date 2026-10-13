import type { EditableElement, EditableLineEnd } from '../editor/EditableElement';
import { A, SP_PR_ORDER, ensureChild } from './xmlEdit';

export const LINE_STYLE_ORDER = [
  'noFill',
  'solidFill',
  'gradFill',
  'pattFill',
  'blipFill',
  'prstDash',
  'custDash',
  'round',
  'bevel',
  'miter',
  'headEnd',
  'tailEnd',
  'extLst',
];

function child(parent: Element, name: string): Element | undefined {
  return Array.from(parent.children).find((element) => element.localName === name);
}

export function patchLineEnd(stroke: Element, name: string, end: EditableLineEnd | null): void {
  if (end === null) {
    child(stroke, name)?.remove();
    return;
  }
  const node = ensureChild(stroke, A, `a:${name}`, LINE_STYLE_ORDER);
  node.setAttribute('type', end.type);
  for (const key of ['w', 'len'] as const) {
    const value = end[key];
    if (value === undefined) node.removeAttribute(key);
    else node.setAttribute(key, value);
  }
}

function patchRadius(properties: Element, element: EditableElement): void {
  const radius = element.cornerRadius;
  const preset = element.shapeType ?? child(properties, 'prstGeom')?.getAttribute('prst');
  if (radius == null || !['rect', 'roundRect'].includes(preset ?? '')) return;
  const existing = child(properties, 'prstGeom') ?? child(properties, 'custGeom');
  let geometry = existing;
  if (!geometry || geometry.localName !== 'prstGeom') {
    geometry = properties.ownerDocument.createElementNS(A, 'a:prstGeom');
    if (existing) existing.replaceWith(geometry);
    else
      properties.insertBefore(
        geometry,
        child(properties, 'xfrm')?.nextSibling ?? properties.firstChild,
      );
  }
  const side = Math.min(Math.abs(element.width), Math.abs(element.height));
  const pixels = Number.isFinite(radius) ? Math.max(0, radius) : 0;
  geometry.setAttribute('prst', pixels > 0 ? 'roundRect' : 'rect');
  const adjustments = ensureChild(geometry, A, 'a:avLst', ['avLst']);
  const oldAdjustment = Array.from(adjustments.children).find(
    (node) => node.getAttribute('name') === 'adj',
  );
  if (pixels === 0) {
    oldAdjustment?.remove();
    return;
  }
  const adjustment = oldAdjustment ?? properties.ownerDocument.createElementNS(A, 'a:gd');
  adjustment.setAttribute('name', 'adj');
  adjustment.setAttribute(
    'fmla',
    `val ${side > 0 ? Math.min(50000, Math.round((pixels / side) * 100000)) : 0}`,
  );
  if (!oldAdjustment) adjustments.appendChild(adjustment);
}

export function patchShapeEditingStyle(
  source: Element,
  element: EditableElement,
  previous: EditableElement | undefined,
): void {
  const properties = child(source, 'spPr');
  if (!properties) return;
  if (
    element.cornerRadius != null &&
    (element.cornerRadius !== previous?.cornerRadius ||
      element.width !== previous?.width ||
      element.height !== previous?.height)
  )
    patchRadius(properties, element);
  for (const name of ['headEnd', 'tailEnd'] as const) {
    const end = element[name];
    if (end === undefined || JSON.stringify(end) === JSON.stringify(previous?.[name])) continue;
    patchLineEnd(ensureChild(properties, A, 'a:ln', SP_PR_ORDER), name, end);
  }
  if (element.lineCap !== undefined && element.lineCap !== previous?.lineCap) {
    ensureChild(properties, A, 'a:ln', SP_PR_ORDER).setAttribute('cap', element.lineCap);
  }
}
