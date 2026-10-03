import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseZip } from '../../../src/parser/ZipParser';
import { buildPresentation } from '../../../src/model/Presentation';
import { resolveRelTarget } from '../../../src/parser/RelParser';
import { arrangeSlides } from '../../../src/writer/deckOps';

const SRC = resolve(__dirname, '../../../docs/example/1-chart-and-complex/source.pptx');
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** Fresh bytes in a plain ArrayBuffer — jsdom and Node keep separate realms. */
function sourceBytes(): Uint8Array {
  const buf = readFileSync(SRC);
  const bytes = new Uint8Array(buf.length);
  bytes.set(buf);
  return bytes;
}

async function load(): Promise<JSZip> {
  return JSZip.loadAsync(sourceBytes());
}

async function xml(zip: JSZip, path: string): Promise<Document> {
  return new DOMParser().parseFromString(await zip.file(path)!.async('string'), 'application/xml');
}

function relsPathOf(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}

async function relTargets(zip: JSZip, part: string, typeSuffix: string): Promise<string[]> {
  const rels = zip.file(relsPathOf(part));
  if (!rels) return [];
  const doc = new DOMParser().parseFromString(await rels.async('string'), 'application/xml');
  const dir = part.slice(0, part.lastIndexOf('/'));
  return Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'))
    .filter((rel) => (rel.getAttribute('Type') ?? '').endsWith(typeSuffix))
    .map((rel) => resolveRelTarget(dir, rel.getAttribute('Target') ?? ''));
}

/** Every internal relationship resolves, every part has a content type, slide ids are unique. */
async function expectConsistentPackage(zip: JSZip): Promise<void> {
  const types = await xml(zip, '[Content_Types].xml');
  const overrides = Array.from(types.getElementsByTagNameNS(CT_NS, 'Override')).map((o) =>
    (o.getAttribute('PartName') ?? '').slice(1),
  );
  const defaults = new Set(
    Array.from(types.getElementsByTagNameNS(CT_NS, 'Default')).map((d) =>
      (d.getAttribute('Extension') ?? '').toLowerCase(),
    ),
  );
  for (const part of overrides) expect(zip.file(part), `override ${part}`).not.toBeNull();

  const files: string[] = [];
  zip.forEach((path, file) => {
    if (!file.dir) files.push(path);
  });
  for (const path of files) {
    if (path === '[Content_Types].xml') continue;
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    expect(overrides.includes(path) || defaults.has(ext), `content type for ${path}`).toBe(true);
    if (!path.endsWith('.rels')) continue;
    const owner = path.replace('_rels/', '').replace(/\.rels$/, '');
    const dir = owner.slice(0, owner.lastIndexOf('/'));
    const doc = await xml(zip, path);
    for (const rel of Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'))) {
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const target = resolveRelTarget(dir, rel.getAttribute('Target') ?? '');
      expect(zip.file(target), `${path} → ${target}`).not.toBeNull();
    }
  }

  const presentation = await xml(zip, 'ppt/presentation.xml');
  const ids = Array.from(presentation.getElementsByTagNameNS(P_NS, 'sldId')).map((e) =>
    Number(e.getAttribute('id')),
  );
  expect(new Set(ids).size).toBe(ids.length);
  for (const id of ids) {
    expect(id).toBeGreaterThanOrEqual(256);
    expect(id).toBeLessThanOrEqual(2147483647);
  }
}

async function reload(zip: JSZip) {
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  return buildPresentation(await parseZip(copy.buffer));
}

describe('arrangeSlides', () => {
  it('copies repeated slides and keeps the package consistent', async () => {
    const zip = await load();

    const paths = await arrangeSlides(zip, [1, 2, 2, 1]);

    expect(paths).toHaveLength(4);
    expect(new Set(paths).size).toBe(4);
    expect(paths.slice(0, 2)).toEqual(['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml']);
    await expectConsistentPackage(zip);
    const pres = await reload(zip);
    expect(pres.slides).toHaveLength(4);
  });

  it('gives a copied slide its own chart and embedded workbook', async () => {
    const zip = await load();

    const [, original, copy] = await arrangeSlides(zip, [1, 2, 2]);

    const [originalChart] = await relTargets(zip, original, '/chart');
    const [copiedChart] = await relTargets(zip, copy, '/chart');
    expect(copiedChart).toBeDefined();
    expect(copiedChart).not.toBe(originalChart);
    const [originalBook] = await relTargets(zip, originalChart, '/package');
    const [copiedBook] = await relTargets(zip, copiedChart, '/package');
    expect(copiedBook).not.toBe(originalBook);
    expect(await zip.file(copiedChart)!.async('string')).toBe(
      await zip.file(originalChart)!.async('string'),
    );
  });

  it('shares pictures and the layout between a slide and its copy, but not the notes', async () => {
    const zip = await load();

    const [original, copy] = await arrangeSlides(zip, [1, 1]);

    expect(await relTargets(zip, copy, '/image')).toEqual(await relTargets(zip, original, '/image'));
    expect(await relTargets(zip, copy, '/slideLayout')).toEqual(
      await relTargets(zip, original, '/slideLayout'),
    );
    expect(await relTargets(zip, original, '/notesSlide')).toHaveLength(1);
    expect(await relTargets(zip, copy, '/notesSlide')).toHaveLength(0);
  });

  it('removes dropped slides with everything only they referenced', async () => {
    const zip = await load();

    const paths = await arrangeSlides(zip, [2]);

    expect(paths).toEqual(['ppt/slides/slide2.xml']);
    expect(zip.file('ppt/slides/slide1.xml')).toBeNull();
    expect(zip.file('ppt/notesSlides/notesSlide1.xml')).toBeNull();
    expect(zip.file('ppt/media/image1.png')).toBeNull();
    expect(zip.file('ppt/charts/chart1.xml')).not.toBeNull();
    await expectConsistentPackage(zip);
    expect((await reload(zip)).slides).toHaveLength(1);
  });

  it('reorders slides and leaves untouched parts byte-identical', async () => {
    const zip = await load();
    const before = await (await load()).file('ppt/slides/slide1.xml')!.async('string');
    const theme = await (await load()).file('ppt/theme/theme1.xml')!.async('string');

    const paths = await arrangeSlides(zip, [2, 1]);

    expect(paths).toEqual(['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml']);
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toBe(before);
    expect(await zip.file('ppt/theme/theme1.xml')!.async('string')).toBe(theme);
    await expectConsistentPackage(zip);
  });

  it('drops section and custom-show lists and updates the slide count', async () => {
    const zip = await load();

    await arrangeSlides(zip, [1, 2, 2]);

    const presentation = await zip.file('ppt/presentation.xml')!.async('string');
    expect(presentation).not.toContain('sectionLst');
    expect(presentation).not.toContain('custShowLst');
    const app = await zip.file('docProps/app.xml')!.async('string');
    expect(app).toContain('<Slides>3</Slides>');
  });

  it('keeps new slide ids inside the OOXML range even when the deck starts near the top', async () => {
    const zip = await load();

    await arrangeSlides(zip, [1, 1, 1, 2, 2, 2]);

    await expectConsistentPackage(zip);
  });

  it('falls back to the lowest free slide id once the top of the range is taken', async () => {
    const zip = await load();
    const presentation = await zip.file('ppt/presentation.xml')!.async('string');
    zip.file(
      'ppt/presentation.xml',
      presentation.replace('<p:sldId id="2147477931"', '<p:sldId id="2147483647"'),
      { createFolders: false },
    );

    await arrangeSlides(zip, [1, 2, 2]);

    const ids = Array.from(
      (await xml(zip, 'ppt/presentation.xml')).getElementsByTagNameNS(P_NS, 'sldId'),
    ).map((e) => Number(e.getAttribute('id')));
    expect(ids[2]).toBe(256);
    await expectConsistentPackage(zip);
  });

  it('rejects positions outside the current slide list', async () => {
    const zip = await load();

    await expect(arrangeSlides(zip, [3])).rejects.toThrow(RangeError);
    await expect(arrangeSlides(zip, [0])).rejects.toThrow(RangeError);
  });
});
