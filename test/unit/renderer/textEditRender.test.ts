/**
 * A text edit must not change how the untouched paragraphs look.
 *
 * Asserted against rendered DOM rather than the model, because the loss was only visible
 * in the layout: rebuilding every paragraph dropped the bullet indent and the paragraph
 * spacing, so an edited body lost its indentation while the text itself looked fine.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseZip } from '../../../src/parser/ZipParser';
import { buildPresentation, type PresentationData } from '../../../src/model/Presentation';
import { renderSlide } from '../../../src/renderer/SlideRenderer';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { applyPlainText, readPlainText } from '../../../src/model/nodes/textEdit';

async function deck(): Promise<PresentationData> {
  const raw = readFileSync(resolve(process.cwd(), 'docs/example/1-chart-and-complex/source.pptx'));
  const bytes = new Uint8Array(raw.length);
  bytes.set(raw);
  return buildPresentation(await parseZip(bytes.buffer.slice(0) as ArrayBuffer));
}

function paragraphLayout(pres: PresentationData, id: string): string[] {
  const handle = renderSlide(pres, pres.slides[0], { onNodeRendered: () => {} });
  const box = handle.element.querySelector(`[data-node-id="${id}"]`) as HTMLElement;
  const out = Array.from(box.querySelectorAll('div')).map((d) => d.getAttribute('style') ?? '');
  handle.dispose();
  return out.filter((s) => /text-align|padding-left|margin-bottom/.test(s));
}

describe('editing a text box keeps the rendered paragraph layout', () => {
  it('leaves the unedited paragraphs styled the same, bullets and spacing included', async () => {
    const pres = await deck();
    const shape = pres.slides[0].nodes.find((n) => n.id === '6') as ShapeNodeData;
    const before = paragraphLayout(pres, shape.id);
    expect(before.some((s) => /padding-left: (?!0px)/.test(s))).toBe(true);

    const lines = readPlainText(shape).split('\n');
    lines[0] = `${lines[0]}X`;
    applyPlainText(shape, lines.join('\n'));

    expect(paragraphLayout(pres, shape.id)).toEqual(before);
  });
});
