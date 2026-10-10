/**
 * Editing text must not cost styles once the deck is written and read back.
 *
 * Checked against a real deck rather than a fixture: the shapes that broke were the ones
 * carrying per-run colour, bullets and paragraph spacing that no hand-written case has.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseZip } from '../../../src/parser/ZipParser';
import { buildPresentation } from '../../../src/model/Presentation';
import { materializeSlideData, type SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { applyPlainText, readPlainText } from '../../../src/model/nodes/textEdit';
import { serializeSlide } from '../../../src/writer/SlideWriter';

/** OOXML markup for style comparison: localName, attributes (no namespace noise), text. */
function fingerprint(el: Element | null | undefined): string {
  if (!el) return '-';
  const attrs = Array.from(el.attributes)
    .filter((a) => !a.name.startsWith('xmlns'))
    .map((a) => `${a.name}=${a.value}`)
    .sort()
    .join(' ');
  const kids = Array.from(el.children).map(fingerprint).join('');
  return `<${el.localName} ${attrs}>${kids || (el.textContent ?? '')}`;
}

/** Paragraph identity: level, its properties, and the paragraph mark. */
function paragraphStyles(shape: ShapeNodeData): string[] {
  return (shape.textBody?.paragraphs ?? []).map(
    (p) => `${p.level}|${fingerprint(p.properties?.element)}|${fingerprint(p.endParaRPr?.element)}`,
  );
}

function runStyles(shape: ShapeNodeData, from: number): string[] {
  return (shape.textBody?.paragraphs ?? [])
    .slice(from)
    .map((p) => p.runs.map((r) => `${r.text}#${fingerprint(r.properties?.element)}`).join(' '));
}

async function load(bytes: Uint8Array) {
  return buildPresentation(await parseZip(bytes));
}

function shapeById(slide: SlideData, id: string): ShapeNodeData {
  return slide.nodes.find((n) => n.id === id) as ShapeNodeData;
}

async function reload(slide: SlideData, id: string): Promise<ShapeNodeData> {
  const raw = readFileSync(resolve(process.cwd(), 'docs/example/1-chart-and-complex/source.pptx'));
  const bytes = new Uint8Array(raw.length);
  bytes.set(raw);
  const { path, xml } = serializeSlide(slide);
  const zip = await JSZip.loadAsync(bytes.buffer.slice(0) as ArrayBuffer);
  zip.file(path, xml, { createFolders: false });
  const composed = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const after = await load(composed);
  materializeSlideData(after.slides[0]);
  const node = after.slides[0].nodes.find((n) => n.id === id) as ShapeNodeData;
  expect(node).toBeTruthy();
  return node;
}

describe('a text edit survives the file', () => {
  it('keeps paragraph and run styles of a real multi-run body', { timeout: 30_000 }, async () => {
    const raw = readFileSync(
      resolve(process.cwd(), 'docs/example/1-chart-and-complex/source.pptx'),
    );
    const bytes = new Uint8Array(raw.length);
    bytes.set(raw);
    const pres = await load(bytes);
    materializeSlideData(pres.slides[0]);
    const shape = pres.slides[0].nodes.find((n) => n.id === '6') as ShapeNodeData;

    const parasBefore = paragraphStyles(shape);
    const runsBefore = runStyles(shape, 1);
    expect(parasBefore.length).toBeGreaterThan(1);

    const lines = readPlainText(shape).split('\n');
    lines[0] = `${lines[0]}X`;
    applyPlainText(shape, lines.join('\n'));
    expect(paragraphStyles(shape)).toEqual(parasBefore);
    expect(runStyles(shape, 1)).toEqual(runsBefore);

    const reloaded = await reload(pres.slides[0], '6');
    expect(paragraphStyles(reloaded)).toEqual(parasBefore);
    expect(runStyles(reloaded, 1)).toEqual(runsBefore);
    expect(readPlainText(reloaded)).toBe(lines.join('\n'));
  });
});
