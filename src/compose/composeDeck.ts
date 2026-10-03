/**
 * Compose a new deck from a template package and a per-page plan.
 *
 * Every page names the template slide it starts from and the edits to make on it. The
 * template's own slides supply everything else — layout, styling, decoration — so a
 * composed page differs from its template only where the plan says so. Pages are laid
 * out first (copied, dropped, reordered as one package operation), then each page's
 * edits are applied to its parsed model, text is fitted, and the slide is written back.
 *
 * Edits that cannot be carried out — an element id the slide does not have, a group the
 * page cannot hold — are reported, not thrown: one bad instruction should cost one edit,
 * not the deck. Only a plan that names a template slide which does not exist is fatal.
 */

import JSZip from 'jszip';
import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import { applyPlainText, readPlainText } from '../model/nodes/textEdit';
import { buildPresentation, materializeSlideNodes } from '../model/Presentation';
import type { SlideData, SlideNode } from '../model/Slide';
import { parseZip } from '../parser/ZipParser';
import { arrangeSlides } from '../writer/deckOps';
import { fitText, type TextFitOptions } from '../writer/fitText';
import { serializeSlide } from '../writer/SlideWriter';
import { layoutItems } from './setItems';
import type { ComposedText, ComposeOp, ComposeProblem, ComposeResult, PagePlan } from './types';

export interface ComposeOptions {
  /** Overrides for text fitting; slide size always comes from the template. */
  fit?: Omit<Partial<TextFitOptions>, 'slideWidth' | 'slideHeight'>;
}

interface PageContext {
  page: number;
  slide: SlideData;
  fit: TextFitOptions;
  texts: ComposedText[];
  problems: ComposeProblem[];
}

function plainBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  return copy.buffer;
}

function isTextShape(node: SlideNode | undefined): node is ShapeNodeData {
  return node?.nodeType === 'shape';
}

function writeText(ctx: PageContext, node: ShapeNodeData, text: string): void {
  applyPlainText(node, text);
  const fitted = fitText(ctx.slide, node, ctx.fit);
  ctx.texts.push({ page: ctx.page, element: node.id, ...fitted });
}

function textTarget(ctx: PageContext, id: string): ShapeNodeData | null {
  const node = ctx.slide.nodes.find((candidate) => candidate.id === id);
  if (isTextShape(node)) return node;
  ctx.problems.push({
    page: ctx.page,
    message: node
      ? `element ${id} is a ${node.nodeType}, not a text shape`
      : `element ${id} is not a top-level element of the slide`,
  });
  return null;
}

function applySetItems(ctx: PageContext, op: Extract<ComposeOp, { op: 'set_items' }>): void {
  const layout = layoutItems(ctx.slide, op.group, op.items.length, {
    width: ctx.fit.slideWidth,
    height: ctx.fit.slideHeight,
  });
  if (typeof layout === 'string') {
    ctx.problems.push({ page: ctx.page, message: layout });
    return;
  }
  for (const note of layout.notes) ctx.problems.push({ page: ctx.page, message: note });
  layout.members.forEach((member, index) => {
    const item = op.items[index];
    for (const [role, node] of member) {
      const text = item[role];
      if (text === undefined) {
        if (isTextShape(node) && readPlainText(node).trim()) {
          ctx.problems.push({
            page: ctx.page,
            message: `item ${index + 1} gives no text for role "${role}"; its template text stays`,
          });
        }
        continue;
      }
      if (isTextShape(node)) writeText(ctx, node, text);
      else ctx.problems.push({ page: ctx.page, message: `role "${role}" is not a text shape` });
    }
  });
}

function applyOp(ctx: PageContext, op: ComposeOp): void {
  if (op.op === 'set_items') {
    applySetItems(ctx, op);
    return;
  }
  const node = textTarget(ctx, op.element);
  if (!node) return;
  if (op.op === 'clear') applyPlainText(node, '');
  else writeText(ctx, node, op.text);
}

/** Build a deck from `template` following `pages`. */
export async function composeDeck(
  template: Uint8Array,
  pages: readonly PagePlan[],
  options: ComposeOptions = {},
): Promise<ComposeResult> {
  const zip = await JSZip.loadAsync(plainBuffer(template));
  await arrangeSlides(
    zip,
    pages.map((page) => page.source),
  );
  const arranged = await zip.generateAsync({ type: 'uint8array' });
  const presentation = buildPresentation(await parseZip(plainBuffer(arranged)));
  const fit: TextFitOptions = {
    ...options.fit,
    slideWidth: presentation.width,
    slideHeight: presentation.height,
  };

  const texts: ComposedText[] = [];
  const problems: ComposeProblem[] = [];
  presentation.slides.forEach((slide, page) => {
    materializeSlideNodes(presentation, slide);
    const ctx: PageContext = { page, slide, fit, texts, problems };
    for (const op of pages[page].ops) applyOp(ctx, op);
    const { path, xml } = serializeSlide(slide);
    zip.file(path, xml, { createFolders: false });
  });

  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes, texts, problems };
}
