/**
 * Fit replacement text into its shape before the slide is written.
 *
 * Swapping a template's sample text for new content leaves the shape sized for the old
 * words. PowerPoint only recomputes an autofit box when someone edits it, and browsers
 * cannot resize an absolutely positioned shape at all, so a longer replacement either
 * spills out of its box or gets squeezed at render time. The writer therefore settles the
 * size here, in model terms, the way an author would:
 *
 *   1. widen the box into free space on its right;
 *   2. grow it into free space below — only where the box is allowed to grow, i.e. its
 *      body says `spAutoFit` (the shape follows its text) or sets no autofit mode;
 *   3. shrink the font, down to a floor.
 *
 * Text metrics are estimated, not measured: average advance widths per script, greedy
 * word wrapping, a fixed line-height factor. That is deliberately cheap and runs without
 * a layout engine. Font sizes inherited from the master's text styles are not resolved
 * here (the caller supplies a default instead). A rendered check is the complement —
 * see `readTextFit` — and catches whatever the estimate misjudges.
 */

import type { ShapeNodeData, TextParagraph, TextRun } from '../model/nodes/ShapeNode';
import type { SlideData, SlideNode } from '../model/Slide';
import { SafeXmlNode } from '../parser/XmlParser';
import { emuToPx } from '../parser/units';
import { A, createEl } from './xmlEdit';

export type TextFitAction = 'fits' | 'widened' | 'grown' | 'shrunk' | 'overflows';

export interface TextFitOptions {
  /** Slide width in pixels. */
  slideWidth: number;
  /** Slide height in pixels. */
  slideHeight: number;
  /** Smallest font scale step 3 may apply. Default 0.7. */
  minScale?: number;
  /** Font size in points for runs whose size is not set in the text body. Default 18. */
  defaultFontSize?: number;
  /** Free margin kept between the box and a neighbour it grows towards, in pixels. Default 8. */
  gutter?: number;
  /** Share of the slide, per edge, a box may not grow into. Default 0.05. */
  pageMargin?: number;
}

export interface TextFitResult {
  action: TextFitAction;
  /** Font scale applied (1 when untouched). */
  scale: number;
}

interface Insets {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

interface Measure {
  width: number;
  height: number;
}

const DEFAULT_INSETS_EMU = { lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720 };
const PX_PER_POINT = 96 / 72;
const CENTI = 100;
const LINE_HEIGHT = 1.2;
const LATIN_EM = 0.55;
const SPACE_EM = 0.28;
const WIDE_EM = 1;
const BOLD_FACTOR = 1.06;
const FIT_TOLERANCE = 1;
const SHRINK_STEPS = 12;
const WIDE_CHAR = /[ᄀ-ᇿ⺀-鿿가-힯豈-﫿＀-￯]/;
const SIZE_STEP = 50;

function bodyAttr(shape: ShapeNodeData, name: string): string | undefined {
  const body = shape.textBody;
  return body?.bodyProperties?.attr(name) ?? body?.layoutBodyProperties?.attr(name);
}

function bodyChild(shape: ShapeNodeData, name: string): boolean {
  const body = shape.textBody;
  const own = body?.bodyProperties?.child(name);
  if (own?.exists()) return true;
  return !!body?.layoutBodyProperties?.child(name).exists();
}

function insetsOf(shape: ShapeNodeData): Insets {
  const read = (name: keyof typeof DEFAULT_INSETS_EMU) =>
    emuToPx(Number(bodyAttr(shape, name) ?? DEFAULT_INSETS_EMU[name]));
  return { left: read('lIns'), top: read('tIns'), right: read('rIns'), bottom: read('bIns') };
}

function listStyleSize(shape: ShapeNodeData, level: number): number | undefined {
  const style = shape.textBody?.listStyle?.child(`lvl${level + 1}pPr`).child('defRPr');
  return style?.numAttr('sz');
}

/** Font size of a run in hundredths of a point, following the text body's own chain. */
function runSize(shape: ShapeNodeData, para: TextParagraph, run: TextRun, fallback: number) {
  return (
    run.properties?.numAttr('sz') ??
    para.properties?.child('defRPr').numAttr('sz') ??
    listStyleSize(shape, para.level) ??
    para.endParaRPr?.numAttr('sz') ??
    fallback * CENTI
  );
}

function lineSpacing(para: TextParagraph): number {
  const pct = para.properties?.child('lnSpc').child('spcPct').numAttr('val');
  return pct ? pct / 100000 : 1;
}

function charWidth(ch: string, sizePx: number, bold: boolean): number {
  const em = ch === ' ' ? SPACE_EM : WIDE_CHAR.test(ch) ? WIDE_EM : LATIN_EM;
  return em * sizePx * (bold ? BOLD_FACTOR : 1);
}

/** Lay one paragraph out greedily at `width`; returns line count and widest line. */
function layoutParagraph(
  pieces: Array<{ ch: string; w: number }>,
  width: number,
): { lines: number; widest: number } {
  let lines = 1;
  let line = 0;
  let widest = 0;
  let word = 0;
  const flushWord = () => {
    if (line > 0 && line + word > width) {
      widest = Math.max(widest, line);
      lines += 1;
      line = 0;
    }
    line += word;
    word = 0;
  };
  for (const { ch, w } of pieces) {
    if (ch === ' ' || WIDE_CHAR.test(ch)) {
      flushWord();
      word = w;
      flushWord();
      continue;
    }
    word += w;
  }
  flushWord();
  return { lines, widest: Math.max(widest, line) };
}

function measure(
  shape: ShapeNodeData,
  innerWidth: number,
  scale: number,
  fallback: number,
): Measure {
  let height = 0;
  let width = 0;
  for (const para of shape.textBody?.paragraphs ?? []) {
    let tallest = 0;
    const pieces: Array<{ ch: string; w: number }> = [];
    for (const run of para.runs) {
      const sizePx = (runSize(shape, para, run, fallback) / CENTI) * PX_PER_POINT * scale;
      const bold = run.properties?.attr('b') === '1';
      tallest = Math.max(tallest, sizePx);
      for (const ch of run.text) pieces.push({ ch, w: charWidth(ch, sizePx, bold) });
    }
    if (!tallest) tallest = fallback * PX_PER_POINT * scale;
    const { lines, widest } = layoutParagraph(pieces, innerWidth);
    height += lines * tallest * LINE_HEIGHT * lineSpacing(para);
    width = Math.max(width, widest);
  }
  return { width, height };
}

function unwrappedWidth(shape: ShapeNodeData, scale: number, fallback: number): number {
  return measure(shape, Number.POSITIVE_INFINITY, scale, fallback).width;
}

function fits(shape: ShapeNodeData, scale: number, fallback: number): boolean {
  const insets = insetsOf(shape);
  const innerWidth = shape.size.w - insets.left - insets.right;
  const innerHeight = shape.size.h - insets.top - insets.bottom;
  const noWrap = bodyAttr(shape, 'wrap') === 'none';
  const laid = measure(shape, noWrap ? Number.POSITIVE_INFINITY : innerWidth, scale, fallback);
  return laid.width <= innerWidth + FIT_TOLERANCE && laid.height <= innerHeight + FIT_TOLERANCE;
}

function overlaps(aStart: number, aEnd: number, bStart: number, bEnd: number): boolean {
  return aStart < bEnd && bStart < aEnd;
}

function roomRight(shape: SlideNode, others: SlideNode[], limit: number, gutter: number): number {
  const { x, y } = shape.position;
  const right = x + shape.size.w;
  let edge = limit;
  for (const other of others) {
    if (other.position.x < right - FIT_TOLERANCE) continue;
    if (!overlaps(y, y + shape.size.h, other.position.y, other.position.y + other.size.h)) continue;
    edge = Math.min(edge, other.position.x - gutter);
  }
  return edge - x;
}

function roomBelow(shape: SlideNode, others: SlideNode[], limit: number, gutter: number): number {
  const { x, y } = shape.position;
  const bottom = y + shape.size.h;
  let edge = limit;
  for (const other of others) {
    if (other.position.y < bottom - FIT_TOLERANCE) continue;
    if (!overlaps(x, x + shape.size.w, other.position.x, other.position.x + other.size.w)) continue;
    edge = Math.min(edge, other.position.y - gutter);
  }
  return edge - y;
}

function canGrow(shape: ShapeNodeData): boolean {
  return bodyChild(shape, 'spAutoFit') || !bodyChild(shape, 'normAutofit');
}

function heightFor(shape: ShapeNodeData, scale: number, fallback: number): number {
  const insets = insetsOf(shape);
  const innerWidth = shape.size.w - insets.left - insets.right;
  return measure(shape, innerWidth, scale, fallback).height + insets.top + insets.bottom;
}

function scaledSize(size: number, scale: number): number {
  return Math.max(SIZE_STEP, Math.round((size * scale) / SIZE_STEP) * SIZE_STEP);
}

function scaleRunSizes(shape: ShapeNodeData, scale: number, fallback: number): void {
  for (const para of shape.textBody?.paragraphs ?? []) {
    for (const run of para.runs) {
      const current = runSize(shape, para, run, fallback);
      ensureRunProperties(shape, run).element?.setAttribute(
        'sz',
        String(scaledSize(current, scale)),
      );
    }
    const end = para.endParaRPr?.numAttr('sz');
    if (end) para.endParaRPr?.element?.setAttribute('sz', String(scaledSize(end, scale)));
  }
}

function ensureRunProperties(shape: ShapeNodeData, run: TextRun): SafeXmlNode {
  if (run.properties?.exists()) return run.properties;
  const doc = shape.source.element!.ownerDocument;
  const props = new SafeXmlNode(createEl(doc, A, 'a:rPr'));
  run.properties = props;
  return props;
}

/** Largest scale in [floor, 1] at which the text fits, or `floor` when nothing does. */
function fittingScale(shape: ShapeNodeData, floor: number, fallback: number): number {
  let low = floor;
  let high = 1;
  if (fits(shape, floor, fallback)) {
    for (let step = 0; step < SHRINK_STEPS; step++) {
      const mid = (low + high) / 2;
      if (fits(shape, mid, fallback)) low = mid;
      else high = mid;
    }
  }
  return low;
}

/**
 * Make `shape`'s text fit its box on `slide`, changing the box size and, as a last
 * resort, the font size in the model. The writer persists both on the next save.
 */
export function fitText(
  slide: SlideData,
  shape: ShapeNodeData,
  options: TextFitOptions,
): TextFitResult {
  const fallback = options.defaultFontSize ?? 18;
  if (!shape.textBody || fits(shape, 1, fallback)) return { action: 'fits', scale: 1 };

  const margin = options.pageMargin ?? 0.05;
  const gutter = options.gutter ?? 8;
  const others = slide.nodes.filter((node) => node !== shape);
  const insets = insetsOf(shape);

  const wanted = unwrappedWidth(shape, 1, fallback) + insets.left + insets.right;
  const room = roomRight(shape, others, options.slideWidth * (1 - margin), gutter);
  const widened = Math.max(shape.size.w, Math.min(room, wanted));
  if (widened > shape.size.w + FIT_TOLERANCE) {
    shape.size = { ...shape.size, w: widened };
    if (fits(shape, 1, fallback)) return { action: 'widened', scale: 1 };
  }

  if (canGrow(shape)) {
    const below = roomBelow(shape, others, options.slideHeight * (1 - margin), gutter);
    const needed = heightFor(shape, 1, fallback);
    const grown = Math.max(shape.size.h, Math.min(below, needed));
    if (grown > shape.size.h + FIT_TOLERANCE) {
      shape.size = { ...shape.size, h: grown };
      if (fits(shape, 1, fallback)) return { action: 'grown', scale: 1 };
    }
  }

  const floor = options.minScale ?? 0.7;
  const scale = fittingScale(shape, floor, fallback);
  scaleRunSizes(shape, scale, fallback);
  return { action: fits(shape, 1, fallback) ? 'shrunk' : 'overflows', scale };
}
