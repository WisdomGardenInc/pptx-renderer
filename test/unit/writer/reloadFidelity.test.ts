/**
 * Serialization must be lossless *as seen by the renderer*, not merely produce a package
 * something can open.
 *
 * The earlier round-trip test checked that python-pptx accepts the output, which it did
 * even in cases where a reload would have lost content — a shape whose element survives
 * but whose box collapsed to zero renders as nothing, and the file still opens fine. So
 * these reload the composed package through the real parser and compare what the model
 * actually ends up holding.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseZip } from '../../../src/parser/ZipParser';
import { buildPresentation } from '../../../src/model/Presentation';
import { materializeSlideData, type SlideData } from '../../../src/model/Slide';
import { serializeSlide } from '../../../src/writer/SlideWriter';

const CASES = ['1-chart-and-complex', 'embedded-font', 'image-crop-css-reset', 'table-stale-frame'];

async function load(bytes: Uint8Array) {
  return buildPresentation(await parseZip(bytes.buffer.slice(0) as ArrayBuffer));
}

function fingerprint(slide: SlideData): string[] {
  return slide.nodes.map(
    (n) => `${n.nodeType}:${n.id}:${Math.round(n.size.w)}x${Math.round(n.size.h)}`,
  );
}

function readCase(stem: string): Uint8Array {
  const buf = readFileSync(resolve(__dirname, '../../../docs/example', stem, 'source.pptx'));
  const bytes = new Uint8Array(buf.length);
  bytes.set(buf);
  return bytes;
}

describe('a serialized slide reloads to the same model', () => {
  for (const stem of CASES) {
    it(`${stem}: node types, ids and boxes are unchanged`, async () => {
      const bytes = readCase(stem);

      const before = await load(bytes);
      const original = before.slides[0];
      materializeSlideData(original);
      const expected = fingerprint(original);
      expect(expected.length).toBeGreaterThan(0);

      // No edit: serialization on its own must change nothing the renderer can see.
      const { path, xml } = serializeSlide(original);

      const zip = await JSZip.loadAsync(bytes);
      zip.file(path, xml, { createFolders: false });
      const composed = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });

      const after = await load(composed);
      const reloaded = after.slides[0];
      materializeSlideData(reloaded);

      expect(fingerprint(reloaded)).toEqual(expected);
    });
  }

  // Decompresses and compares every entry of a 46-part package, which is well past the
  // default budget once the full suite is running in parallel.
  it('keeps every other package entry byte-identical', { timeout: 30_000 }, async () => {
    const bytes = readCase('1-chart-and-complex');
    const pres = await load(bytes);
    materializeSlideData(pres.slides[0]);
    const { path, xml } = serializeSlide(pres.slides[0]);

    const zip = await JSZip.loadAsync(bytes);
    zip.file(path, xml, { createFolders: false });
    const composed = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });

    const before = await JSZip.loadAsync(bytes);
    const after = await JSZip.loadAsync(composed);
    expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());

    for (const name of Object.keys(before.files)) {
      if (name === path || before.files[name].dir) continue;
      const a = await before.files[name].async('uint8array');
      const b = await after.files[name].async('uint8array');
      expect(b, `entry ${name} changed`).toEqual(a);
    }
  });
});
