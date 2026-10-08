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
import type { BaseNodeData } from '../model/nodes/BaseNode';
import type { PicNodeData } from '../model/nodes/PicNode';
import { applyPlainText, readPlainText } from '../model/nodes/textEdit';
import { buildPresentation, materializeSlideNodes } from '../model/Presentation';
import type { SlideData, SlideNode } from '../model/Slide';
import { parseZip } from '../parser/ZipParser';
import { arrangeSlides } from '../writer/deckOps';
import { fitText, fitTextInPlace, type TextFitOptions } from '../writer/fitText';
import { addMediaPart } from '../writer/mediaPart';
import { serializeSlide, slidePartPath } from '../writer/SlideWriter';
import { setPictureImage, syncText } from '../writer/syncNode';
import { layoutItems } from './setItems';
import { templateElements, type TemplateElement, type TemplateElements } from './templateElements';
import type { ComposedText, ComposeOp, ComposeProblem, ComposeResult, PagePlan } from './types';

export interface ComposeOptions {
  /** Overrides for text fitting; slide size always comes from the template. */
  fit?: Omit<Partial<TextFitOptions>, 'slideWidth' | 'slideHeight'>;
}

interface PageContext {
  page: number;
  slide: SlideData;
  zip: JSZip;
  elements: TemplateElements;
  fit: TextFitOptions;
  texts: ComposedText[];
  problems: ComposeProblem[];
}

function plainBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  return copy.buffer;
}

function isTextShape(node: SlideNode | BaseNodeData | undefined): node is ShapeNodeData {
  return node?.nodeType === 'shape';
}

function isPicture(node: SlideNode | BaseNodeData | undefined): node is PicNodeData {
  return node?.nodeType === 'picture';
}

function writeTopLevelText(ctx: PageContext, node: ShapeNodeData, text: string): void {
  applyPlainText(node, text);
  const fitted = fitText(ctx.slide, node, ctx.fit);
  ctx.texts.push({ page: ctx.page, element: node.id, ...fitted });
}

function writeText(ctx: PageContext, entry: TemplateElement, text: string): void {
  const node = entry.node as ShapeNodeData;
  if (entry.parentIds.length === 0) {
    writeTopLevelText(ctx, node, text);
    return;
  }
  // A group child's box is the group's business: the plan may replace its words, not
  // resize it, so its text only shrinks within the box it already has.
  applyPlainText(node, text);
  const fitted = fitTextInPlace(node, ctx.fit);
  syncText(node);
  ctx.texts.push({ page: ctx.page, element: node.id, ...fitted });
}

function textTarget(ctx: PageContext, id: string): TemplateElement | null {
  const entry = ctx.elements.byId.get(id);
  if (entry && isTextShape(entry.node)) return entry;
  ctx.problems.push({
    page: ctx.page,
    message: entry
      ? `element ${id} is a ${entry.node.nodeType}, not a text shape`
      : `element ${id} is not an element of the slide`,
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
      if (isTextShape(node)) writeTopLevelText(ctx, node, text);
      else ctx.problems.push({ page: ctx.page, message: `role "${role}" is not a text shape` });
    }
  });
}

/**
 * The `EG_Media` children a picture can carry, and what each one stands for. A media
 * placeholder keeps its poster in an ordinary blip, so replacing that image would swap the
 * still of a clip while leaving the clip behind.
 */
const MEDIA_POSTERS: Record<string, string | undefined> = {
  audioCd: 'audio',
  wavAudioFile: 'audio',
  audioFile: 'audio',
  videoFile: 'video',
  quickTimeFile: 'video',
};

function posterKind(pic: PicNodeData): string | null {
  const nvPr = pic.source.child('nvPicPr').child('nvPr');
  for (const child of nvPr.children()) {
    const kind = MEDIA_POSTERS[child.localName];
    if (kind) return kind;
  }
  return null;
}

/** Whether `entry` is still a node of the live page, which the ops before this one may have changed. */
function isOnThePage(slide: SlideData, entry: TemplateElement): boolean {
  const outermost = entry.parentIds[0] ?? entry.node.id;
  return slide.nodes.some((node) => node.id === outermost);
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function applySetImage(
  ctx: PageContext,
  op: Extract<ComposeOp, { op: 'set_image' }>,
): Promise<void> {
  const entry = ctx.elements.byId.get(op.element);
  if (!entry) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} is not an element of the slide`,
    });
    return;
  }
  if (!isOnThePage(ctx.slide, entry)) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} is no longer on the page`,
    });
    return;
  }
  if (!isPicture(entry.node)) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} is a ${entry.node.nodeType}, not a picture`,
    });
    return;
  }
  const pic = entry.node;
  const poster = posterKind(pic);
  if (poster) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} is a ${poster} poster, not a picture`,
    });
    return;
  }

  let relId: string;
  try {
    ({ relId } = await addMediaPart(ctx.zip, slidePartPath(ctx.slide), op.bytes));
  } catch (error) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} could not take the new image: ${describeError(error)}`,
    });
    return;
  }
  if (!setPictureImage(pic, relId)) {
    ctx.problems.push({
      page: ctx.page,
      message: `element ${op.element} has no image to replace`,
    });
  }
}

async function applyOp(ctx: PageContext, op: ComposeOp): Promise<void> {
  if (op.op === 'set_items') {
    applySetItems(ctx, op);
    return;
  }
  if (op.op === 'set_image') {
    await applySetImage(ctx, op);
    return;
  }
  const entry = textTarget(ctx, op.element);
  if (!entry) return;
  if (op.op === 'clear') {
    applyPlainText(entry.node as ShapeNodeData, '');
    if (entry.parentIds.length > 0) syncText(entry.node as ShapeNodeData);
  } else {
    writeText(ctx, entry, op.text);
  }
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
  for (const [page, slide] of presentation.slides.entries()) {
    materializeSlideNodes(presentation, slide);
    const ctx: PageContext = {
      page,
      slide,
      zip,
      elements: templateElements(presentation, slide),
      fit,
      texts,
      problems,
    };
    for (const op of pages[page].ops) await applyOp(ctx, op);
    const { path, xml } = serializeSlide(slide);
    zip.file(path, xml, { createFolders: false });
  }

  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes, texts, problems };
}
