/**
 * Media-part writing: put picture bytes into a package and let one part reference them.
 *
 * Compositing replaces a template slide's picture with one drawn for the new deck. The
 * template's own media is left alone: slide copies share their media parts (see `deckOps`),
 * so overwriting `ppt/media/imageN.png` would change every slide that references it, including
 * slides the plan never touched. Each call allocates a fresh part instead, and returns the
 * relationship id for the caller to write into the picture's `r:embed`.
 *
 * The format comes from the bytes rather than from a name the caller supplies, so a part can
 * never be written under an extension its content type contradicts. Content types are declared
 * per extension and nothing else in the writer declares them, so a package that arrives without
 * one for the format gains it here.
 */

import type JSZip from 'jszip';
import {
  directoryOf,
  freshPath,
  readXml,
  relationships,
  relativeTarget,
  relsPathOf,
  writeXml,
} from './deckOps';

const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CONTENT_TYPES = '[Content_Types].xml';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const MEDIA_STEM = 'ppt/media/picture';

interface AddedMedia {
  /** The part that now holds the bytes. */
  path: string;
  /** The relationship the owner part must reference to reach it. */
  relId: string;
}

interface ImageFormat {
  magic: readonly number[];
  extension: string;
  contentType: string;
}

/** The image formats a picture slot can take, told apart by their leading bytes. */
const IMAGE_FORMATS: readonly ImageFormat[] = [
  { magic: [0x89, 0x50, 0x4e, 0x47], extension: 'png', contentType: 'image/png' },
  { magic: [0xff, 0xd8, 0xff], extension: 'jpeg', contentType: 'image/jpeg' },
];

function imageFormat(bytes: Uint8Array): ImageFormat | null {
  return IMAGE_FORMATS.find((entry) => entry.magic.every((byte, i) => bytes[i] === byte)) ?? null;
}

function defaultFor(types: Document, extension: string): Element | null {
  return (
    Array.from(types.getElementsByTagNameNS(CT_NS, 'Default')).find(
      (entry) => (entry.getAttribute('Extension') ?? '').toLowerCase() === extension,
    ) ?? null
  );
}

/** Defaults are declared before overrides, as producers write them. */
function addDefault(types: Document, extension: string, contentType: string): void {
  const element = types.createElementNS(CT_NS, 'Default');
  element.setAttribute('Extension', extension);
  element.setAttribute('ContentType', contentType);
  const firstOverride = types.getElementsByTagNameNS(CT_NS, 'Override')[0];
  if (firstOverride) types.documentElement.insertBefore(element, firstOverride);
  else types.documentElement.appendChild(element);
}

function freeRelId(rels: Document): string {
  const taken = new Set(relationships(rels).map((rel) => rel.id));
  let next = 1;
  while (taken.has(`rId${next}`)) next++;
  return `rId${next}`;
}

/**
 * Store `bytes` as a new media part and point `ownerPath` at it.
 *
 * The part written is a fresh sibling of the template's own media, which leaves every other
 * slide that references that media as it was. Returns the new part's path and the relationship
 * id to embed.
 */
export async function addMediaPart(
  zip: JSZip,
  ownerPath: string,
  bytes: Uint8Array,
): Promise<AddedMedia> {
  const types = await readXml(zip, CONTENT_TYPES);
  const rels = await readXml(zip, relsPathOf(ownerPath));
  if (!types) throw new Error('Not a presentation package');
  if (!rels) throw new Error(`${ownerPath} has no relationships`);

  const image = imageFormat(bytes);
  if (!image) throw new Error('Unsupported image bytes: expected PNG or JPEG');

  const path = freshPath(zip, `${MEDIA_STEM}.${image.extension}`);
  zip.file(path, bytes, { createFolders: false });

  if (!defaultFor(types, image.extension)) addDefault(types, image.extension, image.contentType);
  writeXml(zip, CONTENT_TYPES, types);

  const relId = freeRelId(rels);
  const rel = rels.createElementNS(PKG_REL_NS, 'Relationship');
  rel.setAttribute('Id', relId);
  rel.setAttribute('Type', IMAGE_REL_TYPE);
  rel.setAttribute('Target', relativeTarget(directoryOf(ownerPath), path));
  rels.documentElement.appendChild(rel);
  writeXml(zip, relsPathOf(ownerPath), rels);

  return { path, relId };
}
