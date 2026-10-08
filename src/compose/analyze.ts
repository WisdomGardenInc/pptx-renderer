/**
 * Describe a template deck page by page for whoever plans edits on it.
 *
 * Each page lists its elements — what they are, where, what text they carry and how large
 * — and the repeated groups among them (see `groups.ts`). Elements inside a group are
 * listed too, each naming the groups it sits in, because the text of a card or an agenda
 * item lives there and has to be replaceable like any other. The ids are the ones
 * `composeDeck` acts on, so a plan written against this description can be executed as is.
 * Meaning (which text is the title, which group is a list of KPIs) is not inferred here;
 * it is left to a reader that can see the page.
 */

import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import type { NodeType } from '../model/nodes/BaseNode';
import { readPlainText } from '../model/nodes/textEdit';
import { buildPresentation, materializeSlideNodes } from '../model/Presentation';
import { parseZip, type ZipParseLimits } from '../parser/ZipParser';
import { detectGroups, type DetectedGroup } from './groups';
import { templateElements, type TemplateElement } from './templateElements';

export interface PageElement {
  id: string;
  name: string;
  kind: NodeType;
  x: number;
  y: number;
  w: number;
  h: number;
  /** Plain text, empty for elements without any. */
  text: string;
  /** Largest explicit run size in points, when the text sets one. */
  fontSize?: number;
  /** Placeholder type (`title`, `body`, …) when the element is a placeholder. */
  placeholder?: string;
  /** Ids of the groups the element sits in, outermost first; omitted when top-level. */
  parentIds?: string[];
}

export interface PageAnalysis {
  /** 1-based position in the deck — the `source` a page plan names. */
  index: number;
  elements: PageElement[];
  groups: DetectedGroup[];
}

export interface DeckAnalysis {
  width: number;
  height: number;
  pages: PageAnalysis[];
}

const CENTI = 100;

function fontSizeOf(shape: ShapeNodeData): number | undefined {
  const sizes = (shape.textBody?.paragraphs ?? [])
    .flatMap((para) => para.runs)
    .map((run) => run.properties?.numAttr('sz'))
    .filter((size): size is number => size !== undefined);
  return sizes.length ? Math.max(...sizes) / CENTI : undefined;
}

function describe(entry: TemplateElement): PageElement {
  const node = entry.node;
  const shape = node.nodeType === 'shape' ? (node as ShapeNodeData) : null;
  const element: PageElement = {
    id: node.id,
    name: node.name,
    kind: node.nodeType,
    x: entry.bounds.x,
    y: entry.bounds.y,
    w: entry.bounds.w,
    h: entry.bounds.h,
    text: shape ? readPlainText(shape).trim() : '',
  };
  const fontSize = shape ? fontSizeOf(shape) : undefined;
  if (fontSize !== undefined) element.fontSize = fontSize;
  if (node.placeholder?.type) element.placeholder = node.placeholder.type;
  if (entry.parentIds.length) element.parentIds = entry.parentIds;
  return element;
}

/** Analyse every slide of a PPTX package; `limits` guards against oversized packages. */
export async function analyzeDeck(
  bytes: Uint8Array,
  limits: ZipParseLimits = {},
): Promise<DeckAnalysis> {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  const presentation = buildPresentation(await parseZip(copy.buffer, limits));
  const page = { width: presentation.width, height: presentation.height };
  const pages = presentation.slides.map((slide, position) => {
    materializeSlideNodes(presentation, slide);
    return {
      index: position + 1,
      elements: templateElements(presentation, slide).all.map(describe),
      groups: detectGroups(slide.nodes, page),
    };
  });
  return { ...page, pages };
}
