/**
 * DOM editing primitives for writing OOXML back into a parsed slide part.
 *
 * Two rules drive everything here:
 *
 * 1. **Element order is schema-enforced.** `CT_ShapeProperties` and friends are
 *    `xsd:sequence`, not `xsd:all`, so a fill appended after `a:ln` makes
 *    PowerPoint report the file as corrupt. Every insertion therefore goes
 *    through `ensureChild`, which places the element at its declared slot.
 * 2. **Create elements with their prefix.** `createElementNS(A, 'a:off')` reuses
 *    the `xmlns:a` already declared on the part root; `createElement('a:off')`
 *    would emit an element whose prefix resolves to nothing.
 */

export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Child order of `a:spPr` (CT_ShapeProperties), by localName. */
export const SP_PR_ORDER = [
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
];

/** Child order of `a:xfrm` (CT_Transform2D). */
export const XFRM_ORDER = ['off', 'ext'];

/** The mutually exclusive fill choice in `a:spPr` — setting one clears the rest. */
/**
 * Schema sequence of a:rPr children (CT_TextCharacterProperties). OOXML is a strict sequence,
 * not an unordered set: a fill appended after <a:latin> makes PowerPoint repair the part and
 * drop the styling, so authoring code must insert through this order.
 */
export const RPR_CHILD_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst',
] as const;

export const FILL_LOCAL_NAMES = [
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
];

/** Create a namespaced element, e.g. `createEl(doc, A, 'a:off')`. */
export function createEl(doc: Document, ns: string, qualifiedName: string): Element {
  return doc.createElementNS(ns, qualifiedName);
}

/** First direct child with the given localName, or null. */
export function childByLocal(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.children.length; i++) {
    if (parent.children[i].localName === localName) return parent.children[i];
  }
  return null;
}

/** Remove every direct child whose localName is in `localNames`. */
export function removeChildrenByLocal(parent: Element, localNames: readonly string[]): void {
  const doomed: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    if (localNames.includes(parent.children[i].localName)) doomed.push(parent.children[i]);
  }
  for (const el of doomed) el.remove();
}

/**
 * Insert `el` into `parent` at the slot `order` declares for it, replacing any
 * existing child of the same localName. Children not named in `order` keep their
 * relative position at the end.
 */
export function insertOrdered(parent: Element, el: Element, order: readonly string[]): void {
  const existing = childByLocal(parent, el.localName);
  if (existing) {
    parent.replaceChild(el, existing);
    return;
  }

  const slot = order.indexOf(el.localName);
  if (slot < 0) {
    parent.appendChild(el);
    return;
  }

  for (let i = 0; i < parent.children.length; i++) {
    const rank = order.indexOf(parent.children[i].localName);
    if (rank > slot || rank < 0) {
      parent.insertBefore(el, parent.children[i]);
      return;
    }
  }
  parent.appendChild(el);
}

/**
 * Get the child with `localName`, creating it at its ordered slot when absent.
 */
export function ensureChild(
  parent: Element,
  ns: string,
  qualifiedName: string,
  order: readonly string[],
): Element {
  const localName = qualifiedName.includes(':')
    ? qualifiedName.slice(qualifiedName.indexOf(':') + 1)
    : qualifiedName;
  const existing = childByLocal(parent, localName);
  if (existing) return existing;

  const created = createEl(parent.ownerDocument, ns, qualifiedName);
  insertOrdered(parent, created, order);
  return created;
}

/**
 * Walk up from `el` to the direct child of `root` that contains it, or null when
 * `el` is not a descendant of `root`.
 *
 * Needed because a node's `source` element is not always a direct child of the
 * shape tree: `mc:AlternateContent` wraps its branches, so the element to remove
 * or reorder is the wrapper, not the shape itself.
 */
export function topLevelAncestor(el: Element, root: Element): Element | null {
  let cur: Element | null = el;
  while (cur && cur.parentElement !== root) {
    cur = cur.parentElement;
  }
  return cur;
}

/**
 * Set an attribute, or remove it when `value` is undefined — OOXML treats an
 * absent attribute and its default value as the same thing, and writing the
 * default back adds noise to every diff.
 */
export function setOrRemoveAttr(el: Element, name: string, value: string | undefined): void {
  if (value === undefined) el.removeAttribute(name);
  else el.setAttribute(name, value);
}
