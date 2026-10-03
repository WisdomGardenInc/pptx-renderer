/**
 * Duplicate nodes on a slide — a card, a step, a group with everything inside it.
 *
 * The copy is a deep clone of the source markup, so it keeps every style the original
 * carries, including ones the model does not describe. What a clone cannot keep is its
 * ids: `cNvPr/@id` must be unique across the slide, and that applies to every shape in a
 * group, not only the group itself. All of them are renumbered here, at copy time —
 * the writer's identity sync only revisits top-level ids, and a parsed-looking node whose
 * id matches its markup is left alone even when another shape already owns that id.
 *
 * Connectors inside the copied set (`a:stCxn`, `a:endCxn`) are pointed at the copies of
 * the shapes they join, so a copied step keeps its own arrows instead of reaching back to
 * the original. A chart frame copied this way still refers to the original chart part;
 * duplicating whole slides (`arrangeSlides`) is the route that gives charts their own data.
 */

import { parseRenderableChild } from '../model/RenderableChild';
import type { SlideData, SlideNode } from '../model/Slide';
import { SafeXmlNode } from '../parser/XmlParser';
import { collectUsedIds } from './syncNode';

export interface NodeOffset {
  /** Horizontal shift in pixels. */
  dx: number;
  /** Vertical shift in pixels. */
  dy: number;
}

const CONNECTION_ENDS = ['stCxn', 'endCxn'];

function elementsByLocalName(root: Element, localName: string): Element[] {
  const found: Element[] = [];
  if (root.localName === localName) found.push(root);
  const all = root.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) found.push(all[i]);
  }
  return found;
}

/** Ids already taken on the slide: its markup plus nodes not yet written into it. */
function takenIds(slide: SlideData, spTree: Element): Set<number> {
  const used = collectUsedIds(spTree);
  for (const node of slide.nodes) {
    const el = node.source.element;
    if (!el || spTree.contains(el)) continue;
    collectUsedIds(el).forEach((id) => used.add(id));
  }
  return used;
}

function nextFreeId(used: Set<number>): number {
  let id = 1;
  for (const taken of used) id = Math.max(id, taken + 1);
  used.add(id);
  return id;
}

/**
 * Copy `nodes` on `slide`, shifted by `offset`, and insert the copies into the slide's
 * node list right after the last original, keeping their relative order. Returns the
 * copies in the order given. The slide must be materialized (it needs its parsed root).
 */
export function duplicateNodes(
  slide: SlideData,
  nodes: readonly SlideNode[],
  offset: NodeOffset,
): SlideNode[] {
  const spTree = slide.root?.child('cSld').child('spTree').element;
  if (!spTree) {
    throw new Error(`Cannot duplicate on slide ${slide.index}: the slide is not materialized`);
  }
  const used = takenIds(slide, spTree);
  const renumbered = new Map<string, string>();

  const clones = nodes.map((node) => {
    const source = node.source.element;
    if (!source) throw new Error(`Node ${node.id} has no source markup to copy`);
    const clone = source.cloneNode(true) as Element;
    for (const props of elementsByLocalName(clone, 'cNvPr')) {
      const fresh = String(nextFreeId(used));
      renumbered.set(props.getAttribute('id') ?? '', fresh);
      props.setAttribute('id', fresh);
    }
    return clone;
  });

  for (const clone of clones) {
    for (const end of CONNECTION_ENDS.flatMap((name) => elementsByLocalName(clone, name))) {
      const target = renumbered.get(end.getAttribute('id') ?? '');
      if (target) end.setAttribute('id', target);
    }
  }

  const copies = clones.map((clone, index) => {
    const original = nodes[index];
    const copy = parseRenderableChild(new SafeXmlNode(clone), {
      rels: slide.rels,
      partPath: slide.slidePath,
    });
    if (!copy) throw new Error(`Node ${original.id} could not be parsed back after copying`);
    copy.position = { x: original.position.x + offset.dx, y: original.position.y + offset.dy };
    copy.size = { ...original.size };
    return copy;
  });

  const last = Math.max(...nodes.map((node) => slide.nodes.indexOf(node)));
  const at = last < 0 ? slide.nodes.length : last + 1;
  slide.nodes.splice(at, 0, ...copies);
  return copies;
}
