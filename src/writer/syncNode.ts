/**
 * Push edits made on the typed model back into the XML the node was parsed from.
 *
 * The editor mutates two different things: typed fields (`position`, `size`,
 * `textBody`) and detached `SafeXmlNode` refs (`fill`, `line`, run `properties`).
 * Neither reaches the original tree on its own — the renderer reads the model, not
 * the XML — so this module is what makes `source` tell the truth again before the
 * part is serialized.
 *
 * Detached elements are reused rather than copied: `appendChild` moves a node out of
 * whatever holds it, so the element the renderer reads and the element written to the
 * file stay the same object, and a later edit cannot silently update only one of them.
 */

import type { BaseNodeData } from '../model/nodes/BaseNode';
import type { PicNodeData } from '../model/nodes/PicNode';
import type { ShapeNodeData, TextBody, TextParagraph, TextRun } from '../model/nodes/ShapeNode';
import type { SlideNode } from '../model/Slide';
import { pxToEmu, degToAngle } from '../parser/units';
import {
  A,
  FILL_LOCAL_NAMES,
  SP_PR_ORDER,
  XFRM_ORDER,
  childByLocal,
  createEl,
  ensureChild,
  insertOrdered,
  removeChildrenByLocal,
  setOrRemoveAttr,
} from './xmlEdit';

/** Wrappers that hold a shape's `p:cNvPr`, one per shape kind. */
const NV_PROPS_WRAPPERS = ['nvSpPr', 'nvPicPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr'];

/** Child order of `a:p` (CT_TextParagraph). */
const PARAGRAPH_ORDER = ['pPr', 'r', 'br', 'fld', 'tab', 'endParaRPr'];

/** Run-like children of `a:p`, in the order the model flattens them. */
const RUN_LOCAL_NAMES = ['r', 'br', 'fld', 'tab'];

/**
 * The element holding a node's transform, created when absent.
 *
 * Absence is meaningful rather than exceptional: a placeholder inherits its box from
 * the layout and carries no `a:xfrm` at all until someone moves it, at which point
 * the shape needs one of its own.
 */
function ensureXfrm(el: Element): Element | null {
  const doc = el.ownerDocument;

  // Graphic frames (tables, charts) put the transform directly under p:graphicFrame,
  // in the presentation namespace — not inside an a:spPr like every other shape.
  if (el.localName === 'graphicFrame') {
    const existing = childByLocal(el, 'xfrm');
    if (existing) return existing;
    const created = createEl(doc, el.namespaceURI ?? '', `${el.prefix ?? 'p'}:xfrm`);
    const nvPr = childByLocal(el, 'nvGraphicFramePr');
    if (nvPr?.nextSibling) el.insertBefore(created, nvPr.nextSibling);
    else el.appendChild(created);
    return created;
  }

  const container = childByLocal(el, 'spPr') ?? childByLocal(el, 'grpSpPr');
  if (!container) return null;
  return ensureChild(container, A, 'a:xfrm', SP_PR_ORDER);
}

/** Write `position` / `size` / `rotation` / `flipH` / `flipV` into `a:xfrm`. */
export function syncGeometry(node: BaseNodeData): void {
  const el = node.source.element;
  if (!el) return;

  const xfrm = ensureXfrm(el);
  if (!xfrm) return;

  const off = ensureChild(xfrm, A, 'a:off', XFRM_ORDER);
  off.setAttribute('x', String(pxToEmu(node.position.x)));
  off.setAttribute('y', String(pxToEmu(node.position.y)));

  const ext = ensureChild(xfrm, A, 'a:ext', XFRM_ORDER);
  ext.setAttribute('cx', String(pxToEmu(node.size.w)));
  ext.setAttribute('cy', String(pxToEmu(node.size.h)));

  // Defaults are written as absence: rot="0" and flipH="0" are what a missing
  // attribute already means, so emitting them would only add diff noise.
  setOrRemoveAttr(xfrm, 'rot', node.rotation ? String(degToAngle(node.rotation)) : undefined);
  setOrRemoveAttr(xfrm, 'flipH', node.flipH ? '1' : undefined);
  setOrRemoveAttr(xfrm, 'flipV', node.flipV ? '1' : undefined);
}

/** Move a detached element into `doc` when it came from another parse. */
function adopt(doc: Document, el: Element): Element {
  return el.ownerDocument === doc ? el : (doc.importNode(el, true) as Element);
}

/**
 * Take an element for insertion, cloning it when the model shares it elsewhere.
 *
 * Two runs pointing at one `a:rPr` look identical in the model and are not in the file:
 * `appendChild` moves, so the second owner steals the element and the first paragraph
 * silently loses its font size.
 */
function claim(doc: Document, el: Element, owned: Set<Element>): Element {
  const fresh = owned.has(el) ? adopt(doc, el.cloneNode(true) as Element) : adopt(doc, el);
  owned.add(el);
  return fresh;
}

/** Put `el` into `parent` at its ordered slot, leaving it alone when it is already there. */
function attach(
  parent: Element,
  el: Element | null,
  order: readonly string[],
  owned: Set<Element>,
): void {
  if (!el) return;
  if (childByLocal(parent, el.localName) === el) {
    owned.add(el);
    return;
  }
  insertOrdered(parent, claim(parent.ownerDocument, el, owned), order);
}

/** Write the shape's current fill and line refs into `a:spPr`. */
export function syncShapeStyle(shape: ShapeNodeData): void {
  const el = shape.source.element;
  if (!el) return;
  const spPr = childByLocal(el, 'spPr');
  if (!spPr) return;
  const doc = el.ownerDocument;

  if (shape.fill?.element) {
    // The fill choice is mutually exclusive, so the previous one goes before the
    // new one lands — otherwise a shape ends up with both a solidFill and a gradFill.
    removeChildrenByLocal(spPr, FILL_LOCAL_NAMES);
    insertOrdered(spPr, adopt(doc, shape.fill.element), SP_PR_ORDER);
  }

  if (shape.line?.element) {
    insertOrdered(spPr, adopt(doc, shape.line.element), SP_PR_ORDER);
  }
}

/** Build the element for one run, reusing its `a:rPr` and, for fields, the original element. */
function buildRunElement(
  doc: Document,
  run: TextRun,
  original: Element | undefined,
  owned: Set<Element>,
): Element {
  const rPr = run.properties?.element ?? null;

  // a:fld carries a generated GUID in @id that must survive the round trip, so the
  // original element is reused whenever the run still is a field.
  if (run.fieldType && original?.localName === 'fld') {
    const tEl = ensureChild(original, A, 'a:t', ['rPr', 'pPr', 't']);
    tEl.textContent = run.text;
    attach(original, rPr, ['rPr', 'pPr', 't'], owned);
    return original;
  }

  if (run.text === '\n') {
    const br = createEl(doc, A, 'a:br');
    attach(br, rPr, ['rPr'], owned);
    return br;
  }

  if (run.text === '\t') {
    return createEl(doc, A, 'a:tab');
  }

  const r = createEl(doc, A, 'a:r');
  attach(r, rPr, ['rPr', 't'], owned);
  const t = createEl(doc, A, 'a:t');
  t.textContent = run.text;
  r.appendChild(t);
  return r;
}

/** Rewrite one paragraph's run children from the model. */
function syncParagraphRuns(pEl: Element, para: TextParagraph, owned: Set<Element>): void {
  const doc = pEl.ownerDocument;

  const originals: Element[] = [];
  for (let i = 0; i < pEl.children.length; i++) {
    if (RUN_LOCAL_NAMES.includes(pEl.children[i].localName)) originals.push(pEl.children[i]);
  }

  // Built before anything is removed: constructing a run moves its rPr (and, for a
  // field, the whole element) out of the old child, so the old children must still
  // be attached at this point.
  const built = para.runs.map((run, i) => buildRunElement(doc, run, originals[i], owned));

  for (const old of originals) old.remove();

  const endParaRPr = childByLocal(pEl, 'endParaRPr');
  for (const el of built) {
    if (endParaRPr) pEl.insertBefore(el, endParaRPr);
    else pEl.appendChild(el);
  }
}

/** Rewrite a shape's `a:txBody` paragraphs from the model. */
export function syncText(shape: ShapeNodeData): void {
  const el = shape.source.element;
  if (!el || !shape.textBody) return;
  const txBody = childByLocal(el, 'txBody');
  if (!txBody) return;
  const doc = el.ownerDocument;

  const paras = shape.textBody.paragraphs;

  const parsed: Element[] = [];
  for (let i = 0; i < txBody.children.length; i++) {
    if (txBody.children[i].localName === 'p') parsed.push(txBody.children[i]);
  }

  // Match model paragraphs to their a:p. A paragraph parsed from this part still points
  // at its a:pPr, whose parent is the a:p to claim. The rest — ones the editor created,
  // ones imported from another Document — take the free a:p elements in order, so a
  // paragraph nobody rewrote keeps the markup it was parsed with.
  const target: (Element | null)[] = paras.map(() => null);
  const claimed: Set<Element> = new Set();
  paras.forEach((para, i) => {
    const pPrEl = para.properties?.element;
    const p = pPrEl && pPrEl.parentElement?.parentElement === txBody ? pPrEl.parentElement : null;
    if (p && !claimed.has(p)) {
      target[i] = p;
      claimed.add(p);
    }
  });
  const free = parsed.filter((p) => !claimed.has(p));
  let next = 0;
  target.forEach((t, i) => {
    if (t || next >= free.length) return;
    target[i] = free[next];
    claimed.add(free[next]);
    next += 1;
  });

  const owned = new Set<Element>();
  const kept: Element[] = [];
  paras.forEach((para, i) => {
    const pEl = target[i] ?? createEl(doc, A, 'a:p');
    kept.push(pEl);

    // The model owns paragraph properties once the text is rewritten, and a paragraph
    // cloned by an edit carries a detached pPr: it has to be attached here or the
    // alignment, bullets and spacing the edit preserved never reach the file.
    attach(pEl, para.properties?.element ?? null, PARAGRAPH_ORDER, owned);

    if (para.level > 0) {
      const pPr = ensureChild(pEl, A, 'a:pPr', PARAGRAPH_ORDER);
      pPr.setAttribute('lvl', String(para.level));
    }

    syncParagraphRuns(pEl, para, owned);

    attach(pEl, para.endParaRPr?.element ?? null, PARAGRAPH_ORDER, owned);
  });

  // Drop paragraphs the edit removed, then reattach in model order. Appending an
  // already-attached element moves it, so this doubles as the reorder.
  const doomed: Element[] = [];
  for (let i = 0; i < txBody.children.length; i++) {
    const child = txBody.children[i];
    if (child.localName === 'p' && !kept.includes(child)) doomed.push(child);
  }
  for (const p of doomed) p.remove();
  for (const p of kept) txBody.appendChild(p);
}

/** A node's `p:cNvPr`, whichever non-visual wrapper its shape kind uses. */
function findCNvPr(el: Element): Element | null {
  for (const wrapper of NV_PROPS_WRAPPERS) {
    const found = childByLocal(el, wrapper);
    if (found) return childByLocal(found, 'cNvPr');
  }
  return childByLocal(el, 'cNvPr');
}

/** Remove the vector alternatives a blip may carry, which viewers prefer over its raster one. */
function dropVectorBlips(blip: Element): void {
  const extLst = childByLocal(blip, 'extLst');
  if (!extLst) return;
  for (const ext of Array.from(extLst.children)) {
    if (childByLocal(ext, 'svgBlip')) ext.remove();
  }
}

/**
 * Point a picture at a new image relationship, reporting whether it could be done.
 *
 * The artwork it used to show goes with it. A linked blip keeps its target outside the
 * package, and a vector blip in `a:blip/extLst` is preferred over the raster one — either
 * left behind keeps showing the template's own picture. The crop goes too: it was cut for
 * the image being replaced, and the replacement arrives already sized for the slot.
 *
 * Unlike the other sync functions this one is driven by an edit the caller holds rather than
 * by the model, so it is called for the picture being replaced and not from `syncNode` —
 * nothing else on the slide should be rewritten just because a picture elsewhere changed.
 */
export function setPictureImage(pic: PicNodeData, relId: string): boolean {
  const el = pic.source.element;
  if (!el) return false;
  const blipFill = childByLocal(el, 'blipFill');
  const blip = blipFill ? childByLocal(blipFill, 'blip') : null;
  if (!blipFill || !blip) return false;

  blip.setAttribute('r:embed', relId);
  blip.removeAttribute('embed');
  blip.removeAttribute('link');
  blip.removeAttribute('r:link');
  dropVectorBlips(blip);
  removeChildrenByLocal(blipFill, ['srcRect']);

  pic.blipEmbed = relId;
  pic.blipLink = undefined;
  pic.crop = undefined;
  return true;
}
/**
 * Write the node's id and name into `p:cNvPr`.
 *
 * For a parsed node this is a no-op — the id came from that very attribute. It matters
 * for nodes the editor created: those are built from a seed XML string with a fixed
 * placeholder id, and only their typed `id` was ever stamped. Without this, every added
 * shape lands in the file under the same id, and on reload they collide in the host's
 * node map so that all but the last become unselectable.
 *
 * `cNvPr/@id` is `xsd:unsignedInt` and must be unique within the slide, so a node whose
 * id is not a positive integer (an early `new-…` id, say) is renumbered here rather than
 * written out as invalid markup. The model is updated to match, since the caller holds
 * that id too.
 */
function syncIdentity(node: SlideNode, used: Set<number>): void {
  const el = node.source.element;
  if (!el) return;
  const cNvPr = findCNvPr(el);
  if (!cNvPr) return;

  const current = cNvPr.getAttribute('id');
  const isValid = (value: string | null): boolean => {
    const n = Number(value);
    return value !== null && value !== '' && Number.isInteger(n) && n >= 1;
  };

  // A parsed node that nobody renumbered already agrees with its markup.
  if (node.id === current && isValid(current)) {
    if (node.name) cNvPr.setAttribute('name', node.name);
    return;
  }

  let numeric = Number(node.id);
  if (!isValid(node.id) || used.has(numeric)) {
    numeric = 1;
    while (used.has(numeric)) numeric += 1;
    node.id = String(numeric);
  }
  used.add(numeric);
  cNvPr.setAttribute('id', String(numeric));
  if (node.name) cNvPr.setAttribute('name', node.name);
}

/** Every `cNvPr/@id` in the part, including ids no model node covers. */
export function collectUsedIds(root: Element): Set<number> {
  const used = new Set<number>();
  const all = root.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName !== 'cNvPr') continue;
    const value = Number(all[i].getAttribute('id'));
    if (Number.isInteger(value) && value >= 1) used.add(value);
  }
  return used;
}

/** Push every supported edit on one node back into its XML. */
export function syncNode(node: SlideNode, usedIds?: Set<number>): void {
  if (usedIds) syncIdentity(node, usedIds);
  syncGeometry(node);
  if (node.nodeType === 'shape') {
    const shape = node as ShapeNodeData;
    syncShapeStyle(shape);
    if (shape.textBody) syncText(shape);
  }
}

/** Exported for tests: whether a text body carries anything worth writing. */
export function hasParagraphs(body: TextBody | undefined): boolean {
  return !!body && body.paragraphs.length > 0;
}
