/**
 * Describe a template deck page by page for whoever plans edits on it.
 *
 * Each page lists its top-level elements — what they are, where, what text they carry and
 * how large — and the repeated groups among them (see `groups.ts`). The ids are the ones
 * `composeDeck` acts on, so a plan written against this description can be executed as is.
 * Meaning (which text is the title, which group is a list of KPIs) is not inferred here;
 * it is left to a reader that can see the page.
 */

import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import { readPlainText } from '../model/nodes/textEdit';
import { buildPresentation, materializeSlideNodes } from '../model/Presentation';
import type { SlideNode } from '../model/Slide';
import { parseZip, type ZipParseLimits } from '../parser/ZipParser';
import { detectGroups, type DetectedGroup } from './groups';

export interface PageElement {
  id: string;
  name: string;
  kind: SlideNode['nodeType'];
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

function describe(node: SlideNode): PageElement {
  const shape = node.nodeType === 'shape' ? (node as ShapeNodeData) : null;
  const element: PageElement = {
    id: node.id,
    name: node.name,
    kind: node.nodeType,
    x: node.position.x,
    y: node.position.y,
    w: node.size.w,
    h: node.size.h,
    text: shape ? readPlainText(shape).trim() : '',
  };
  const fontSize = shape ? fontSizeOf(shape) : undefined;
  if (fontSize !== undefined) element.fontSize = fontSize;
  if (node.placeholder?.type) element.placeholder = node.placeholder.type;
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
      elements: slide.nodes.map(describe),
      groups: detectGroups(slide.nodes, page),
    };
  });
  return { ...page, pages };
}
