import type { EditableElement } from '../editor/EditableElement';
import { remapGroupGeometry } from '../model/groupGeometry';
import { inheritedGroupFill } from '../model/groupFill';
import { parseOoxmlBool } from '../parser/booleans';

export type GroupAction = 'group' | 'ungroup';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function child(parent: Element, name: string): Element | undefined {
  return Array.from(parent.children).find((element) => element.localName === name);
}

function nodeId(node: Element): string | null {
  const first = node.children[0];
  return first ? (child(first, 'cNvPr')?.getAttribute('id') ?? null) : null;
}

function transform(node: Element): Element {
  const properties = child(node, node.localName === 'grpSp' ? 'grpSpPr' : 'spPr');
  const xfrm =
    node.localName === 'graphicFrame'
      ? child(node, 'xfrm')
      : properties && child(properties, 'xfrm');
  if (!xfrm) throw new Error('The selected object has no editable transform.');
  return xfrm;
}

function vector(parent: Element, name: string, x: string, y: string): [number, number] {
  const element = child(parent, name);
  if (!element) throw new Error('The selected object has no editable geometry.');
  return [Number(element.getAttribute(x)), Number(element.getAttribute(y))];
}

function setVector(
  parent: Element,
  name: string,
  x: string,
  y: string,
  first: number,
  second: number,
): void {
  const element = child(parent, name);
  if (!element) throw new Error('The selected object has no editable geometry.');
  element.setAttribute(x, String(Math.round(first)));
  element.setAttribute(y, String(Math.round(second)));
}

function findNode(document: Document, id: string): Element {
  const property = Array.from(document.getElementsByTagNameNS(P, 'cNvPr')).find(
    (element) => element.getAttribute('id') === id,
  );
  const node = property?.parentElement?.parentElement;
  if (!node) throw new Error('The selected object no longer exists.');
  return node;
}

function appendVector(
  document: Document,
  parent: Element,
  name: string,
  x: string,
  y: string,
  first: number,
  second: number,
): void {
  const element = document.createElementNS(A, `a:${name}`);
  element.setAttribute(x, String(Math.round(first)));
  element.setAttribute(y, String(Math.round(second)));
  parent.appendChild(element);
}

function groupNodes(document: Document, ids: readonly string[]): string {
  const nodes = ids.map((id) => findNode(document, id));
  const parent = nodes[0].parentElement;
  if (!parent || nodes.length < 2 || nodes.some((node) => node.parentElement !== parent))
    throw new Error('Select objects in the same group.');
  const ordered = Array.from(parent.children).filter((node) => nodes.includes(node));
  const boxes = ordered.map((node) => {
    const xfrm = transform(node);
    const [x, y] = vector(xfrm, 'off', 'x', 'y');
    const [w, h] = vector(xfrm, 'ext', 'cx', 'cy');
    const angle = ((Number(xfrm.getAttribute('rot') ?? 0) / 60000) * Math.PI) / 180;
    const rotatedWidth = Math.abs(w * Math.cos(angle)) + Math.abs(h * Math.sin(angle));
    const rotatedHeight = Math.abs(w * Math.sin(angle)) + Math.abs(h * Math.cos(angle));
    return {
      x: x + (w - rotatedWidth) / 2,
      y: y + (h - rotatedHeight) / 2,
      w: rotatedWidth,
      h: rotatedHeight,
    };
  });
  const left = Math.min(...boxes.map((box) => box.x));
  const top = Math.min(...boxes.map((box) => box.y));
  const width = Math.max(1, Math.max(...boxes.map((box) => box.x + box.w)) - left);
  const height = Math.max(1, Math.max(...boxes.map((box) => box.y + box.h)) - top);
  const id = String(
    Math.max(
      ...Array.from(document.getElementsByTagNameNS(P, 'cNvPr')).map((element) =>
        Number(element.getAttribute('id')),
      ),
    ) + 1,
  );
  const group = document.createElementNS(P, 'p:grpSp');
  const nv = document.createElementNS(P, 'p:nvGrpSpPr');
  const properties = document.createElementNS(P, 'p:cNvPr');
  properties.setAttribute('id', id);
  properties.setAttribute('name', `Group ${id}`);
  nv.append(
    properties,
    document.createElementNS(P, 'p:cNvGrpSpPr'),
    document.createElementNS(P, 'p:nvPr'),
  );
  const groupProperties = document.createElementNS(P, 'p:grpSpPr');
  const xfrm = document.createElementNS(A, 'a:xfrm');
  appendVector(document, xfrm, 'off', 'x', 'y', left, top);
  appendVector(document, xfrm, 'ext', 'cx', 'cy', width, height);
  appendVector(document, xfrm, 'chOff', 'x', 'y', left, top);
  appendVector(document, xfrm, 'chExt', 'cx', 'cy', width, height);
  groupProperties.appendChild(xfrm);
  group.append(nv, groupProperties);
  parent.insertBefore(group, ordered[ordered.length - 1].nextSibling);
  for (const node of ordered) group.appendChild(node);
  return id;
}

function ungroupNode(document: Document, id: string): string[] {
  const group = findNode(document, id);
  const parent = group.parentElement;
  if (group.localName !== 'grpSp' || !parent)
    throw new Error('The selected object is not a group.');
  const gx = transform(group);
  const [ox, oy] = vector(gx, 'off', 'x', 'y');
  const [gw, gh] = vector(gx, 'ext', 'cx', 'cy');
  const [cx, cy] = vector(gx, 'chOff', 'x', 'y');
  const [cw, ch] = vector(gx, 'chExt', 'cx', 'cy');
  const groupGeometry = {
    size: { w: gw, h: gh },
    childOffset: { x: cx, y: cy },
    childExtent: { w: cw, h: ch },
    flipH: parseOoxmlBool(gx.getAttribute('flipH') ?? undefined),
    flipV: parseOoxmlBool(gx.getAttribute('flipV') ?? undefined),
  };
  const rotation = Number(gx.getAttribute('rot') ?? 0) / 60000;
  const radians = (rotation * Math.PI) / 180;
  const members = Array.from(group.children).filter(
    (element) => !['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(element.localName),
  );
  if (members.some((element) => !nodeId(element)))
    throw new Error('The group contains an object that cannot be ungrouped.');
  const children = members;
  for (const node of children) {
    const xfrm = transform(node);
    const [x, y] = vector(xfrm, 'off', 'x', 'y');
    const [w, h] = vector(xfrm, 'ext', 'cx', 'cy');
    const angle = Number(xfrm.getAttribute('rot') ?? 0) / 60000;
    const geometry = {
      position: { x, y },
      size: { w, h },
      rotation: angle,
      flipH: parseOoxmlBool(xfrm.getAttribute('flipH') ?? undefined),
      flipV: parseOoxmlBool(xfrm.getAttribute('flipV') ?? undefined),
    };
    remapGroupGeometry(geometry, groupGeometry, node.localName !== 'graphicFrame');
    const dx = geometry.position.x + geometry.size.w / 2 - gw / 2;
    const dy = geometry.position.y + geometry.size.h / 2 - gh / 2;
    const nextX =
      ox + gw / 2 + dx * Math.cos(radians) - dy * Math.sin(radians) - geometry.size.w / 2;
    const nextY =
      oy + gh / 2 + dx * Math.sin(radians) + dy * Math.cos(radians) - geometry.size.h / 2;
    setVector(xfrm, 'off', 'x', 'y', nextX, nextY);
    setVector(xfrm, 'ext', 'cx', 'cy', geometry.size.w, geometry.size.h);
    xfrm.setAttribute('rot', String(Math.round((geometry.rotation + rotation) * 60000)));
    if (node.localName !== 'graphicFrame') {
      if (groupGeometry.flipH) xfrm.setAttribute('flipH', geometry.flipH ? '1' : '0');
      if (groupGeometry.flipV) xfrm.setAttribute('flipV', geometry.flipV ? '1' : '0');
    }
    const properties = child(node, node.localName === 'grpSp' ? 'grpSpPr' : 'spPr');
    const groupFill = properties && child(properties, 'grpFill');
    if (groupFill) {
      const inherited = inheritedGroupFill(node);
      const replacement = inherited?.cloneNode(true) ?? document.createElementNS(A, 'a:noFill');
      groupFill.replaceWith(replacement);
    }
    parent.insertBefore(node, group);
  }
  group.remove();
  return children.map((node) => nodeId(node)!);
}

export function restoreUngroupedGeometry(
  elements: readonly EditableElement[],
  previous: readonly EditableElement[],
  nodeIds: readonly string[],
): EditableElement[] {
  const previousById = new Map(previous.map((element) => [element.id, element]));
  const changedNodes = new Set(nodeIds);
  return elements.map((element) => {
    if (!element.nodeId || !changedNodes.has(element.nodeId)) return element;
    const original = previousById.get(element.id);
    if (!original) return element;
    return {
      ...element,
      x: original.x,
      y: original.y,
      width: original.width,
      height: original.height,
      rotation: original.rotation,
      flipH: original.flipH,
      flipV: original.flipV,
    };
  });
}

export function editGroups(
  xml: string,
  action: GroupAction,
  ids: readonly string[],
): { xml: string; ids: string[] } {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const result = action === 'group' ? [groupNodes(document, ids)] : ungroupNode(document, ids[0]);
  return { xml: new XMLSerializer().serializeToString(document), ids: result };
}
