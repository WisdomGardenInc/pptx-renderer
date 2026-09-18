import {
  SafeXmlNode,
  type ShapeNodeData,
  type TextRun,
} from '@wisdomgarden/pptx-renderer';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export interface TextStyle {
  fontSize?: number; // points
  fontFamily?: string;
  bold: boolean;
  italic: boolean;
  color?: string; // hex, no '#'
}

export interface TextStylePatch {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

/**
 * Reads editable style from a shape's model refs, and writes it back by mutating / replacing
 * those same refs (run.properties, node.fill). The renderer reads these typed refs directly,
 * so a re-render reflects the change — no OOXML re-serialization, and no staleness from cached
 * elements (we reassign the ref rather than editing a detached subtree).
 */

function ownerDoc(shape: ShapeNodeData): Document {
  return (
    shape.source?.element?.ownerDocument ??
    document.implementation.createDocument(DML, 'a:root', null)
  );
}

function el(doc: Document, local: string): Element {
  return doc.createElementNS(DML, `a:${local}`);
}

function removeChildrenByLocal(parent: Element, local: string): void {
  for (const c of Array.from(parent.children)) {
    if (c.localName === local) c.remove();
  }
}

function setSolidColor(parent: Element, hex: string, doc: Document): void {
  removeChildrenByLocal(parent, 'solidFill');
  const sf = el(doc, 'solidFill');
  const clr = el(doc, 'srgbClr');
  clr.setAttribute('val', hex.replace('#', '').toUpperCase());
  sf.appendChild(clr);
  parent.appendChild(sf);
}

/** All non-empty runs across the shape's paragraphs. */
function runs(shape: ShapeNodeData): TextRun[] {
  return shape.textBody?.paragraphs.flatMap((p) => p.runs) ?? [];
}

function rPrElement(run: TextRun, doc: Document): Element {
  const existing = run.properties?.element ?? null;
  if (existing) return existing;
  const created = el(doc, 'rPr');
  run.properties = new SafeXmlNode(created);
  return created;
}

// ---- read ------------------------------------------------------------------

export function readTextStyle(shape: ShapeNodeData): TextStyle {
  const first = runs(shape).find((r) => r.text && r.text.trim().length > 0) ?? runs(shape)[0];
  const rPr = first?.properties;
  const szRaw = rPr?.attr('sz');
  return {
    fontSize: szRaw ? Number(szRaw) / 100 : undefined,
    fontFamily: rPr?.child('latin').attr('typeface'),
    bold: rPr?.attr('b') === '1' || rPr?.attr('b') === 'true',
    italic: rPr?.attr('i') === '1' || rPr?.attr('i') === 'true',
    color: rPr?.child('solidFill').child('srgbClr').attr('val'),
  };
}

export function readFillColor(shape: ShapeNodeData): string | null {
  const fill = shape.fill;
  if (fill && fill.exists() && fill.localName === 'solidFill') {
    return fill.child('srgbClr').attr('val') ?? null;
  }
  return null;
}

export function readPlainText(shape: ShapeNodeData): string {
  return (
    shape.textBody?.paragraphs
      .map((p) => p.runs.map((r) => r.text).join(''))
      .join('\n') ?? ''
  );
}

// ---- write -----------------------------------------------------------------

export function applyTextStyle(shape: ShapeNodeData, patch: TextStylePatch): void {
  const doc = ownerDoc(shape);
  for (const run of runs(shape)) {
    const rPr = rPrElement(run, doc);
    if (patch.fontSize !== undefined) rPr.setAttribute('sz', String(Math.round(patch.fontSize * 100)));
    if (patch.bold !== undefined) {
      if (patch.bold) rPr.setAttribute('b', '1');
      else rPr.removeAttribute('b');
    }
    if (patch.italic !== undefined) {
      if (patch.italic) rPr.setAttribute('i', '1');
      else rPr.removeAttribute('i');
    }
    if (patch.fontFamily !== undefined) {
      let latin = Array.from(rPr.children).find((c) => c.localName === 'latin');
      if (!latin) {
        latin = el(doc, 'latin');
        rPr.appendChild(latin);
      }
      latin.setAttribute('typeface', patch.fontFamily);
    }
    if (patch.color !== undefined) setSolidColor(rPr, patch.color, doc);
    // Reassign so runs that had no rPr (undefined) now expose the fresh one to the renderer.
    run.properties = new SafeXmlNode(rPr);
  }
}

export function applyFillColor(shape: ShapeNodeData, hex: string): void {
  const doc = ownerDoc(shape);
  const sf = el(doc, 'solidFill');
  const clr = el(doc, 'srgbClr');
  clr.setAttribute('val', hex.replace('#', '').toUpperCase());
  sf.appendChild(clr);
  shape.fill = new SafeXmlNode(sf);
}

/**
 * Replace the shape's text with `text` (newlines → paragraphs), reusing the first run's style
 * so typography survives the edit. Operates on the typed model (renderer reads run.text).
 */
export function setPlainText(shape: ShapeNodeData, text: string): void {
  if (!shape.textBody) {
    shape.textBody = { paragraphs: [] };
  }
  const styleRun = runs(shape).find((r) => r.properties) ?? runs(shape)[0];
  const styleProps = styleRun?.properties;
  const level = shape.textBody.paragraphs[0]?.level ?? 0;
  shape.textBody.paragraphs = text.split('\n').map((line) => ({
    level,
    runs: [{ text: line, properties: styleProps }],
  }));
}
