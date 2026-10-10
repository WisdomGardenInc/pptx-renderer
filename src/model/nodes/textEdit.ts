/**
 * Replace a shape's text the way an inline editor sees it, without discarding the
 * styling around it.
 *
 * A plain-text host — a textarea, a contenteditable region, a properties field — sees a
 * text body as paragraphs joined by newlines. Rebuilding the body from that string alone
 * is what makes styles vanish on the first commit: paragraph properties (alignment,
 * indents, bullets, spacing), the paragraph-mark properties, and the per-run styling of
 * everything the user never touched.
 *
 * So this keeps the structure the text came out of and changes only what the new string
 * actually differs in:
 *
 * - a line whose text is unchanged keeps its paragraph object, and with it its a:pPr,
 *   a:endParaRPr, level, and every run reference;
 * - a line whose text changed keeps the runs at its unchanged head and tail, so only the
 *   edited span is replaced, and the replacement carries the style of the run that used
 *   to sit there;
 * - a paragraph that has to serve two output lines is deep-cloned, never shared: two runs
 *   pointing at one a:rPr means the writer moves that element out of the first one, and
 *   the saved file silently loses a font size.
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import type { ShapeNodeData, TextParagraph, TextRun } from './ShapeNode';

/** Text of one paragraph as a plain-text host sees it; a soft break (a:br) reads as \n. */
export function paragraphPlainText(para: TextParagraph): string {
  return para.runs.map((run) => run.text).join('');
}

/** The whole body as a host's textarea shows it. Round-trips with applyPlainText. */
export function readPlainText(shape: ShapeNodeData): string {
  return shape.textBody?.paragraphs.map(paragraphPlainText).join('\n') ?? '';
}

/** The style-bearing parts of a paragraph, captured before anything is rewritten. */
interface ParagraphStyle {
  properties?: SafeXmlNode;
  level: number;
  endParaRPr?: SafeXmlNode;
  runs: TextRun[];
}

function cloneProps(node: SafeXmlNode | undefined): SafeXmlNode | undefined {
  const el = node?.element;
  return el ? new SafeXmlNode(el.cloneNode(true) as Element) : undefined;
}

/** A detached paragraph carrying the style of `src` and the runs it can be diffed against. */
function cloneParagraph(src: ParagraphStyle): TextParagraph {
  return {
    properties: cloneProps(src.properties),
    level: src.level,
    endParaRPr: cloneProps(src.endParaRPr),
    runs: src.runs.map((run) => ({
      text: run.text,
      fieldType: run.fieldType,
      properties: cloneProps(run.properties),
    })),
  };
}

/**
 * Give one paragraph new text, keeping the runs the change does not reach.
 *
 * The replacement run takes the properties of the run the edit landed on: an overwritten
 * span keeps the style it replaced, and a pure insertion inherits the style to its left,
 * which is what typing into the middle of a word does. A run that survives the edit keeps
 * its own a:rPr, so that style has to be cloned, not moved.
 */
function rewriteParagraphText(para: TextParagraph, text: string): void {
  const before = para.runs.map((r) => r.text).join('');
  if (before === text) return;

  let head = 0;
  while (head < before.length && head < text.length && before[head] === text[head]) head += 1;
  let tail = 0;
  while (
    tail < before.length - head &&
    tail < text.length - head &&
    before[before.length - 1 - tail] === text[text.length - 1 - tail]
  ) {
    tail += 1;
  }

  const spans: { run: TextRun; start: number; end: number }[] = [];
  let pos = 0;
  for (const run of para.runs) {
    spans.push({ run, start: pos, end: pos + run.text.length });
    pos += run.text.length;
  }

  // A run may straddle the unchanged region; then it goes, and the characters it held are
  // re-emitted from the new text. Keeping is therefore decided by whole runs, not by the
  // character trim, or the tail of a replaced run would vanish.
  const keepHead: typeof spans = [];
  for (const span of spans) {
    if (span.end > head) break;
    keepHead.push(span);
  }
  const keepTail: typeof spans = [];
  for (let i = spans.length - 1; i >= 0; i--) {
    if (i < keepHead.length || spans[i].start < before.length - tail) break;
    keepTail.unshift(spans[i]);
  }
  const cutStart = keepHead.length ? keepHead[keepHead.length - 1].end : 0;
  const cutEnd = keepTail.length ? keepTail[0].start : before.length;
  const replaced = spans.slice(keepHead.length, spans.length - keepTail.length);

  const donor = replaced[0] ?? keepHead[keepHead.length - 1] ?? keepTail[0];
  const inserted = text.slice(cutStart, text.length - (before.length - cutEnd));
  const runs: TextRun[] = keepHead.map((s) => s.run);
  if (inserted !== '') {
    runs.push({
      text: inserted,
      // A replaced run is leaving the model, so its element is free to move into the new
      // run. A kept one still owns its a:rPr, which the copy must not share.
      properties: donor?.run.properties
        ? replaced.length > 0
          ? donor.run.properties
          : cloneProps(donor.run.properties)
        : undefined,
    });
  }
  for (const span of keepTail) runs.push(span.run);

  para.runs = runs;
}

/**
 * Replace the shape's text with `text`, reading `\n` as a paragraph break, and keep as
 * much of the surrounding styling as the change allows.
 *
 * Lines are matched against the paragraphs they came from, longest unchanged head and
 * tail first, so an edit in one paragraph never touches the others. Whatever has to be
 * written is written at run level, against the text that paragraph actually held.
 */
export function applyPlainText(shape: ShapeNodeData, text: string): void {
  if (!shape.textBody) shape.textBody = { paragraphs: [] };
  const paras = shape.textBody.paragraphs;
  const lines = text.split('\n');

  if (paras.length === 0) {
    shape.textBody.paragraphs = lines.map((line) => ({ level: 0, runs: [{ text: line }] }));
    return;
  }

  // A soft break makes one paragraph cover several of the host's lines, so alignment
  // happens over flattened lines that still remember which paragraph they came from.
  const oldLines: { text: string; para: number }[] = [];
  paras.forEach((para, i) => {
    for (const seg of paragraphPlainText(para).split('\n')) {
      oldLines.push({ text: seg, para: i });
    }
  });

  // Captured up front: a paragraph serving two output lines is cloned from what it was,
  // not from whatever its first line left behind.
  const styles: ParagraphStyle[] = paras.map((p) => ({
    properties: p.properties,
    level: p.level,
    endParaRPr: p.endParaRPr,
    runs: p.runs.map((r) => ({
      text: r.text,
      fieldType: r.fieldType,
      properties: r.properties,
    })),
  }));

  let head = 0;
  while (head < oldLines.length && head < lines.length && oldLines[head].text === lines[head]) {
    head += 1;
  }
  let tail = 0;
  while (
    tail < oldLines.length - head &&
    tail < lines.length - head &&
    oldLines[oldLines.length - 1 - tail].text === lines[lines.length - 1 - tail]
  ) {
    tail += 1;
  }

  const owner: number[] = new Array(lines.length).fill(0);
  for (let i = 0; i < head; i++) owner[i] = oldLines[i].para;
  for (let i = 0; i < tail; i++) {
    owner[lines.length - 1 - i] = oldLines[oldLines.length - 1 - i].para;
  }
  const midLen = oldLines.length - tail - head;
  const fallback = head > 0 ? oldLines[head - 1].para : oldLines[0].para;
  for (let i = head; i < lines.length - tail; i++) {
    // Extra lines land on the last paragraph of the changed region: a new line inherits
    // the style of the line it was split off, which is what Enter does in PowerPoint.
    owner[i] = midLen > 0 ? oldLines[head + Math.min(i - head, midLen - 1)].para : fallback;
  }

  const out: TextParagraph[] = [];
  const used = new Set<number>();
  for (let i = 0; i < lines.length; ) {
    const para = owner[i];
    let j = i + 1;
    while (j < lines.length && owner[j] === para) j += 1;
    const group = lines.slice(i, j);
    const original = paras[para];

    if (!used.has(para) && group.join('\n') === paragraphPlainText(original)) {
      // Untouched paragraph: keep the object, and with it every run, soft break and
      // element reference the writer is going to reuse.
      out.push(original);
      used.add(para);
    } else {
      group.forEach((line, k) => {
        const target = k === 0 && !used.has(para) ? original : cloneParagraph(styles[para]);
        used.add(para);
        rewriteParagraphText(target, line);
        out.push(target);
      });
    }
    i = j;
  }

  shape.textBody.paragraphs = out;
}
