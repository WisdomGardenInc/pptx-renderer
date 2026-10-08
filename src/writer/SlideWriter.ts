/**
 * Serialize an edited slide back to OOXML.
 *
 * This is an incremental rewrite, not a generator: the part keeps every element it
 * was parsed with, and only what the editor actually touched is changed. Anything
 * the model does not describe — charts, SmartArt, embedded objects, effects the
 * renderer ignores — is carried through untouched simply by never being visited.
 *
 * The output is one slide part. Assembling a package from a baseline plus a set of
 * these is the caller's job, and deliberately so: that step is pure file handling
 * with no OOXML knowledge in it, and it belongs wherever the baseline lives.
 */

import type { SlideData, SlideNode } from '../model/Slide';
import { SafeXmlNode } from '../parser/XmlParser';
import { collectUsedIds, syncNode } from './syncNode';
import { topLevelAncestor } from './xmlEdit';

/** Children of `p:spTree` that are structure rather than content. */
const SP_TREE_FIXED = ['nvGrpSpPr', 'grpSpPr'];

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * Serialize a document to a well-formed OOXML part.
 *
 * `XMLSerializer` does not agree with itself across environments: browsers emit the
 * document's XML declaration, jsdom does not. Prepending one unconditionally produces a
 * file with two declarations — which is not merely untidy, it is fatally malformed, and
 * a slide part that no parser will read renders as an empty slide rather than an error.
 */
export function serializeDocument(doc: Document): string {
  const serialized = new XMLSerializer().serializeToString(doc);
  return serialized.startsWith('<?xml') ? serialized : XML_DECLARATION + serialized;
}

export interface SerializedSlidePart {
  /** Part path inside the package, e.g. `ppt/slides/slide3.xml`. */
  path: string;
  /** Slide XML, declaration included. */
  xml: string;
}

/**
 * The part a slide lives in. Everything that addresses the slide's own package parts — its
 * relationships above all — has to name the same path the part is written to.
 */
export function slidePartPath(slide: SlideData): string {
  return slide.slidePath || `ppt/slides/slide${slide.index + 1}.xml`;
}

/** Locate `p:cSld > p:spTree` under a slide root. */
function findShapeTree(root: SafeXmlNode): Element | null {
  return root.child('cSld').child('spTree').element;
}

/**
 * Make `spTree` hold exactly the nodes the model lists, in model order.
 *
 * A node created by the editor was parsed in its own Document, so it is imported
 * here and its `source` repointed at the imported element — otherwise the next edit
 * would mutate a detached copy and never reach the file.
 *
 * Elements are matched through `topLevelAncestor` because a shape is not always a
 * direct child of the tree: `mc:AlternateContent` wraps its branches, and it is the
 * wrapper that gets moved or dropped.
 */
function alignShapeTree(spTree: Element, nodes: SlideNode[]): void {
  const doc = spTree.ownerDocument;
  const keep: Element[] = [];

  for (const node of nodes) {
    const el = node.source.element;
    if (!el) continue;

    const attached = topLevelAncestor(el, spTree);
    if (attached) {
      if (!keep.includes(attached)) keep.push(attached);
      continue;
    }

    const imported = doc.importNode(el, true) as Element;
    node.source = new SafeXmlNode(imported);
    spTree.appendChild(imported);
    keep.push(imported);
  }

  const doomed: Element[] = [];
  for (let i = 0; i < spTree.children.length; i++) {
    const child = spTree.children[i];
    if (SP_TREE_FIXED.includes(child.localName)) continue;
    if (!keep.includes(child)) doomed.push(child);
  }
  for (const el of doomed) el.remove();

  // Appending an attached element moves it, so this reorders in place.
  for (const el of keep) spTree.appendChild(el);
}

/**
 * Write the model's edits into the slide's XML tree.
 *
 * Exposed separately from serialization because a host that batches several edits
 * before saving wants the sync without paying for the string.
 */
export function syncSlide(slide: SlideData): void {
  if (!slide.root) return;
  const spTree = findShapeTree(slide.root);
  if (!spTree) return;

  alignShapeTree(spTree, slide.nodes);

  // Ids are allocated against everything the part already declares, not just what the
  // model lists: group children and anything the parser skipped still occupy the slide's
  // id space, and `cNvPr/@id` has to be unique across all of it.
  const usedIds = collectUsedIds(spTree);
  for (const node of slide.nodes) syncNode(node, usedIds);
}

/**
 * Sync the model into the slide's XML and serialize the part.
 *
 * Throws when the slide carries no root element — a lazily parsed slide must be
 * materialized before it can be written.
 */
export function serializeSlide(slide: SlideData): SerializedSlidePart {
  if (!slide.root?.element) {
    throw new Error(
      `Cannot serialize slide ${slide.index}: no parsed root. ` +
        'Materialize the slide (render it, or call materializeSlideData) first.',
    );
  }

  syncSlide(slide);

  const doc = slide.root.element.ownerDocument;

  return {
    path: slidePartPath(slide),
    xml: serializeDocument(doc),
  };
}
