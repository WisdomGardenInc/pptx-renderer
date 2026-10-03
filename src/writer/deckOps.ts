/**
 * Slide-level package operations: build a deck's slide list from an existing package.
 *
 * `arrangeSlides` takes the package's current slides and lays out a new list from them —
 * any subset, in any order, with repeats. A repeat is a real copy: the slide part and its
 * relationships are duplicated, and parts that belong to one slide alone (charts, SmartArt
 * data) are copied with it so that editing one copy never reaches into another. Pictures,
 * layouts and other shared targets stay shared.
 *
 * Everything the new list no longer reaches — dropped slides, their notes, media only they
 * used — is removed afterwards by walking the relationship graph from the package root, so
 * the result never carries orphaned parts or dangling content-type overrides.
 *
 * Section and custom-show lists are dropped when the slide list changes: both name slides
 * by id, and keeping them consistent with an arbitrary rearrangement is not possible in
 * general. PowerPoint treats a deck without them as an ordinary unsectioned deck.
 */

import type JSZip from 'jszip';
import { resolveRelTarget } from '../parser/RelParser';
import { serializeDocument } from './SlideWriter';

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

const CONTENT_TYPES = '[Content_Types].xml';
const PRESENTATION = 'ppt/presentation.xml';
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const APP_PROPS = 'docProps/app.xml';
const ROOT_RELS = '_rels/.rels';
const SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const MIN_SLIDE_ID = 256;
const MAX_SLIDE_ID = 2147483647;

/** Relationship types whose targets travel with a copied slide instead of being shared. */
const OWNED_REL_SUFFIXES = [
  '/chart',
  '/chartEx',
  '/diagramData',
  '/diagramLayout',
  '/diagramQuickStyle',
  '/diagramColors',
  '/diagramDrawing',
];

/** Relationship types a copied slide does not take with it. */
const DROPPED_REL_SUFFIXES = ['/notesSlide'];

/** Relationship types inside an owned part that stay shared (pictures, media). */
const SHARED_REL_SUFFIXES = ['/image', '/media', '/video', '/audio', '/hyperlink'];

/** `p:extLst` children that list slides by id; dropped when the slide list changes. */
const SLIDE_LIST_EXT_URIS = new Set(['{521415D9-36F7-43E2-AB2F-B90AF26B5E84}']);

interface Relationship {
  element: Element;
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function parse(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

async function readXml(zip: JSZip, path: string): Promise<Document | null> {
  const file = zip.file(path);
  return file ? parse(await file.async('string')) : null;
}

function writeXml(zip: JSZip, path: string, doc: Document): void {
  zip.file(path, serializeDocument(doc), { createFolders: false });
}

function directoryOf(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash < 0 ? '' : path.slice(0, slash);
}

function relsPathOf(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}

function relationships(doc: Document): Relationship[] {
  return Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship')).map((element) => ({
    element,
    id: element.getAttribute('Id') ?? '',
    type: element.getAttribute('Type') ?? '',
    target: element.getAttribute('Target') ?? '',
    external: element.getAttribute('TargetMode') === 'External',
  }));
}

function hasSuffix(type: string, suffixes: readonly string[]): boolean {
  return suffixes.some((suffix) => type.endsWith(suffix));
}

/** Next free path beside `path`, numbered after every sibling that shares its stem. */
function freshPath(zip: JSZip, path: string): string {
  const dir = directoryOf(path);
  const name = path.slice(dir.length + 1);
  const match = /^(.*?)(\d*)(\.[^.]+)$/.exec(name);
  const stem = match ? match[1] : name;
  const ext = match ? match[3] : '';
  let highest = 0;
  zip.forEach((relative) => {
    if (directoryOf(relative) !== dir) return;
    const sibling = /^(.*?)(\d*)(\.[^.]+)$/.exec(relative.slice(dir.length + 1));
    if (sibling && sibling[1] === stem && sibling[3] === ext) {
      highest = Math.max(highest, Number(sibling[2] || 0));
    }
  });
  return `${dir}/${stem}${highest + 1}${ext}`;
}

function overrideFor(types: Document, path: string): Element | null {
  const partName = `/${path}`;
  return (
    Array.from(types.getElementsByTagNameNS(CT_NS, 'Override')).find(
      (override) => override.getAttribute('PartName') === partName,
    ) ?? null
  );
}

function addOverride(types: Document, path: string, contentType: string): void {
  const override = types.createElementNS(CT_NS, 'Override');
  override.setAttribute('PartName', `/${path}`);
  override.setAttribute('ContentType', contentType);
  types.documentElement.appendChild(override);
}

/**
 * Copy a part, its relationships and every owned target beneath it. Returns the new path.
 * Copies land in the source's directory, so relative targets stay valid unchanged except
 * where they point at a part that was itself copied.
 */
async function copyPart(
  zip: JSZip,
  types: Document,
  path: string,
  owned: (type: string) => boolean,
): Promise<string> {
  const copy = freshPath(zip, path);
  const bytes = await zip.file(path)!.async('uint8array');
  zip.file(copy, bytes, { createFolders: false });
  const override = overrideFor(types, path);
  if (override) addOverride(types, copy, override.getAttribute('ContentType') ?? '');

  const rels = await readXml(zip, relsPathOf(path));
  if (!rels) return copy;
  for (const rel of relationships(rels)) {
    if (rel.external) continue;
    if (hasSuffix(rel.type, DROPPED_REL_SUFFIXES)) {
      rel.element.parentNode?.removeChild(rel.element);
      continue;
    }
    if (!owned(rel.type)) continue;
    const target = resolveRelTarget(directoryOf(path), rel.target);
    if (!zip.file(target)) continue;
    const targetCopy = await copyPart(
      zip,
      types,
      target,
      (type) => !hasSuffix(type, SHARED_REL_SUFFIXES),
    );
    rel.element.setAttribute('Target', relativeTarget(directoryOf(copy), targetCopy));
  }
  writeXml(zip, relsPathOf(copy), rels);
  return copy;
}

/** Relative target from `fromDir` to `path`, in the `../dir/name` form OOXML uses. */
function relativeTarget(fromDir: string, path: string): string {
  const from = fromDir.split('/').filter(Boolean);
  const to = path.split('/').filter(Boolean);
  let shared = 0;
  while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) shared++;
  return [...from.slice(shared).map(() => '..'), ...to.slice(shared)].join('/');
}

/** Parts reachable from the package root through the relationship graph. */
async function reachableParts(zip: JSZip): Promise<Set<string>> {
  const reached = new Set<string>();
  const pending: Array<{ source: string; rels: string }> = [{ source: '', rels: ROOT_RELS }];
  while (pending.length) {
    const { source, rels: relsPath } = pending.pop()!;
    const rels = await readXml(zip, relsPath);
    if (!rels) continue;
    for (const rel of relationships(rels)) {
      if (rel.external) continue;
      const target = resolveRelTarget(directoryOf(source), rel.target);
      if (reached.has(target) || !zip.file(target)) continue;
      reached.add(target);
      pending.push({ source: target, rels: relsPathOf(target) });
    }
  }
  return reached;
}

/** Remove every part the relationship graph no longer reaches, with its rels and override. */
async function pruneUnreachable(zip: JSZip, types: Document): Promise<void> {
  const reached = await reachableParts(zip);
  const candidates: string[] = [];
  zip.forEach((path, file) => {
    if (file.dir || path === CONTENT_TYPES || path.endsWith('.rels')) return;
    if (!reached.has(path)) candidates.push(path);
  });
  for (const path of candidates) {
    zip.remove(path);
    zip.remove(relsPathOf(path));
    overrideFor(types, path)?.remove();
  }
}

function dropSlideListExtensions(presentation: Document): void {
  for (const ext of Array.from(presentation.getElementsByTagNameNS(P_NS, 'ext'))) {
    if (SLIDE_LIST_EXT_URIS.has(ext.getAttribute('uri') ?? '')) ext.parentNode?.removeChild(ext);
  }
  for (const shows of Array.from(presentation.getElementsByTagNameNS(P_NS, 'custShowLst'))) {
    shows.parentNode?.removeChild(shows);
  }
}

/**
 * A slide id the list does not use yet. Ids run up to 2^31 - 1; decks written by PowerPoint
 * often start near the top of that range, so counting up can run out — fall back to the
 * lowest free id when it does.
 */
function allocateSlideId(taken: Set<number>): number {
  const highest = Math.max(MIN_SLIDE_ID - 1, ...taken);
  let id = highest < MAX_SLIDE_ID ? highest + 1 : MIN_SLIDE_ID;
  while (taken.has(id)) id++;
  taken.add(id);
  return id;
}

async function updateSlideCount(zip: JSZip, count: number): Promise<void> {
  const app = await readXml(zip, APP_PROPS);
  const slides = app?.getElementsByTagName('Slides')[0];
  if (!app || !slides) return;
  slides.textContent = String(count);
  writeXml(zip, APP_PROPS, app);
}

/**
 * Lay out a new slide list from the package's current slides.
 *
 * `order` lists 1-based positions in the current slide list. A position may repeat — each
 * repeat after the first becomes an independent copy — and positions left out are removed
 * together with whatever only they referenced. Returns the slide part paths of the new
 * list, in order.
 */
export async function arrangeSlides(zip: JSZip, order: readonly number[]): Promise<string[]> {
  const presentation = await readXml(zip, PRESENTATION);
  const presRels = await readXml(zip, PRESENTATION_RELS);
  const types = await readXml(zip, CONTENT_TYPES);
  if (!presentation || !presRels || !types) throw new Error('Not a presentation package');

  const list = presentation.getElementsByTagNameNS(P_NS, 'sldIdLst')[0];
  const entries = list ? Array.from(list.getElementsByTagNameNS(P_NS, 'sldId')) : [];
  const relById = new Map(relationships(presRels).map((rel) => [rel.id, rel]));
  const pathOf = (entry: Element) =>
    resolveRelTarget('ppt', relById.get(entry.getAttributeNS(R_NS, 'id') ?? '')?.target ?? '');
  for (const position of order) {
    if (!Number.isInteger(position) || position < 1 || position > entries.length) {
      throw new RangeError(`Slide ${position} is outside 1..${entries.length}`);
    }
  }

  const takenIds = new Set(entries.map((entry) => Number(entry.getAttribute('id'))));
  let nextRel = relById.size + 1;
  const used = new Set<number>();
  const arranged: Array<{ entry: Element; path: string }> = [];

  for (const position of order) {
    const source = entries[position - 1];
    if (!used.has(position)) {
      used.add(position);
      arranged.push({ entry: source, path: pathOf(source) });
      continue;
    }
    const path = await copyPart(zip, types, pathOf(source), (type) =>
      hasSuffix(type, OWNED_REL_SUFFIXES),
    );
    if (!overrideFor(types, path)) addOverride(types, path, SLIDE_CONTENT_TYPE);
    while (relById.has(`rId${nextRel}`)) nextRel++;
    const relId = `rId${nextRel++}`;
    const rel = presRels.createElementNS(PKG_REL_NS, 'Relationship');
    rel.setAttribute('Id', relId);
    rel.setAttribute('Type', SLIDE_REL_TYPE);
    const target = relativeTarget('ppt', path);
    rel.setAttribute('Target', target);
    presRels.documentElement.appendChild(rel);
    relById.set(relId, { element: rel, id: relId, type: SLIDE_REL_TYPE, target, external: false });
    const entry = presentation.createElementNS(P_NS, 'p:sldId');
    entry.setAttribute('id', String(allocateSlideId(takenIds)));
    entry.setAttributeNS(R_NS, 'r:id', relId);
    arranged.push({ entry, path });
  }

  entries.forEach((entry, index) => {
    if (used.has(index + 1)) return;
    relById.get(entry.getAttributeNS(R_NS, 'id') ?? '')?.element.remove();
  });
  if (list) {
    while (list.firstChild) list.removeChild(list.firstChild);
    for (const { entry } of arranged) list.appendChild(entry);
  }
  dropSlideListExtensions(presentation);

  writeXml(zip, PRESENTATION, presentation);
  writeXml(zip, PRESENTATION_RELS, presRels);
  await pruneUnreachable(zip, types);
  writeXml(zip, CONTENT_TYPES, types);
  await updateSlideCount(zip, arranged.length);
  return arranged.map(({ path }) => path);
}
