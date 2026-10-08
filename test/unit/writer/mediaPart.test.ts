import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { resolveRelTarget } from '../../../src/parser/RelParser';
import { addMediaPart } from '../../../src/writer/mediaPart';
import { serializeDocument } from '../../../src/writer/SlideWriter';

const SRC = resolve(__dirname, '../../../docs/example/1-chart-and-complex/source.pptx');
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const OWNER = 'ppt/slides/slide1.xml';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2]);

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

async function relsOf(zip: JSZip, owner: string): Promise<Element[]> {
  const path = owner.replace(/([^/]+)$/, '_rels/$1.rels');
  const doc = await xml(zip, path);
  return Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'));
}

async function defaults(zip: JSZip): Promise<Element[]> {
  const types = await xml(zip, '[Content_Types].xml');
  return Array.from(types.getElementsByTagNameNS(CT_NS, 'Default'));
}

async function partBytes(zip: JSZip, path: string): Promise<Uint8Array> {
  return zip.file(path)!.async('uint8array');
}

/** The relationship ids the owner part already declares. */
async function relIds(zip: JSZip, owner: string): Promise<string[]> {
  return (await relsOf(zip, owner)).map((entry) => entry.getAttribute('Id')!);
}

/** The media parts the package carries, by name. */
function mediaNames(zip: JSZip): string[] {
  return Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/'));
}

/** The sample deck with `path` already occupied by `bytes`. */
async function withPart(path: string, bytes: Uint8Array): Promise<JSZip> {
  const zip = await load();
  zip.file(path, bytes, { createFolders: false });
  return zip;
}

/** The sample deck without its png default, so a new one has to be declared. */
async function withoutPngDefault(): Promise<JSZip> {
  const zip = await load();
  const types = await xml(zip, '[Content_Types].xml');
  for (const entry of Array.from(types.getElementsByTagNameNS(CT_NS, 'Default'))) {
    if (entry.getAttribute('Extension') === 'png') entry.remove();
  }
  zip.file('[Content_Types].xml', serializeDocument(types), { createFolders: false });
  return zip;
}

describe('addMediaPart', () => {
  it('stores a new picture beside the template media and points the owner at it', async () => {
    const zip = await load();
    const before = await relIds(zip, OWNER);

    const { path, relId } = await addMediaPart(zip, OWNER, PNG);

    expect(path).toMatch(/^ppt\/media\/picture\d+\.png$/);
    expect(await partBytes(zip, path)).toEqual(PNG);
    expect(before).not.toContain(relId);
    const rel = (await relsOf(zip, OWNER)).find((entry) => entry.getAttribute('Id') === relId)!;
    expect(rel.getAttribute('Type')).toMatch(/\/image$/);
    expect(resolveRelTarget('ppt/slides', rel.getAttribute('Target')!)).toBe(path);
  });

  it('tells the formats apart by their bytes', async () => {
    const zip = await load();

    const { path } = await addMediaPart(zip, OWNER, JPEG);

    expect(path).toMatch(/^ppt\/media\/picture\d+\.jpeg$/);
    const jpeg = (await defaults(zip)).filter((d) => d.getAttribute('Extension') === 'jpeg');
    expect(jpeg).toHaveLength(1);
    expect(jpeg[0]!.getAttribute('ContentType')).toBe('image/jpeg');
  });

  it('leaves the media the template already had where it is', async () => {
    const zip = await load();
    const before = await partBytes(zip, 'ppt/media/image1.png');
    const other = await partBytes(zip, 'ppt/media/image2.png');

    const { path } = await addMediaPart(zip, OWNER, PNG);

    expect(path).not.toBe('ppt/media/image1.png');
    expect(await partBytes(zip, 'ppt/media/image1.png')).toEqual(before);
    expect(await partBytes(zip, 'ppt/media/image2.png')).toEqual(other);
  });

  it('takes a free name when the template already has that sibling', async () => {
    const taken = new Uint8Array([...PNG, 7]);
    const zip = await withPart('ppt/media/picture1.png', taken);

    const { path } = await addMediaPart(zip, OWNER, PNG);

    expect(path).not.toBe('ppt/media/picture1.png');
    expect(await partBytes(zip, 'ppt/media/picture1.png')).toEqual(taken);
    expect(await partBytes(zip, path)).toEqual(PNG);
  });

  it('declares the extension once, and only when the package lacks it', async () => {
    const kept = await load();
    await addMediaPart(kept, OWNER, PNG);
    const png = (await defaults(kept)).filter((d) => d.getAttribute('Extension') === 'png');
    expect(png).toHaveLength(1);

    const missing = await withoutPngDefault();
    expect((await defaults(missing)).some((d) => d.getAttribute('Extension') === 'png')).toBe(
      false,
    );

    await addMediaPart(missing, OWNER, PNG);

    const types = await xml(missing, '[Content_Types].xml');
    const children = Array.from(types.documentElement.children);
    const added = (await defaults(missing)).find((d) => d.getAttribute('Extension') === 'png')!;
    expect(added.getAttribute('ContentType')).toBe('image/png');
    expect(children.indexOf(added)).toBeLessThan(
      children.findIndex((child) => child.localName === 'Override'),
    );
  });

  it('refuses bytes it cannot type, without writing the part', async () => {
    const zip = await load();
    const before = mediaNames(zip);

    await expect(
      addMediaPart(zip, OWNER, new Uint8Array([0x47, 0x49, 0x46, 0x38])),
    ).rejects.toThrow('Unsupported image bytes: expected PNG or JPEG');

    expect(mediaNames(zip)).toEqual(before);
    expect(await relsOf(zip, OWNER)).toHaveLength(4);
  });

  it('refuses an owner that is not in the package', async () => {
    const zip = await load();

    await expect(addMediaPart(zip, 'ppt/slides/slide9.xml', PNG)).rejects.toThrow(
      'ppt/slides/slide9.xml has no relationships',
    );
  });

  it('gives a second picture its own part and its own relationship', async () => {
    const zip = await load();
    const before = await relIds(zip, OWNER);

    const first = await addMediaPart(zip, OWNER, PNG);
    const second = await addMediaPart(zip, OWNER, JPEG);

    expect(second.relId).not.toBe(first.relId);
    expect(before).not.toContain(second.relId);
    expect(second.path).not.toBe(first.path);
    expect(await partBytes(zip, first.path)).toEqual(PNG);
    expect(await partBytes(zip, second.path)).toEqual(JPEG);
  });
});
