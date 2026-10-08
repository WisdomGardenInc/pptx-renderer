import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import type { GroupSpec } from '../../src/compose/types';
import { buildPresentation, materializeSlideNodes } from '../../src/model/Presentation';
import type { PresentationData } from '../../src/model/Presentation';
import { parseZip } from '../../src/parser/ZipParser';

const SRC = resolve(__dirname, '../../docs/example/1-chart-and-complex/source.pptx');
export const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export const EMU = 9525;
export const STEP = 299;
export const LABELS = ['Quote', 'Apply', 'Underwrite', 'Activate'];

export function xfrm(x: number, y: number, w: number, h: number): string {
  return `<a:xfrm><a:off x="${Math.round(x * EMU)}" y="${Math.round(y * EMU)}"/><a:ext cx="${Math.round(w * EMU)}" cy="${Math.round(h * EMU)}"/></a:xfrm>`;
}

export function textShape(
  id: number,
  name: string,
  box: [number, number, number, number],
  text: string,
  sz = 1400,
) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:r><a:rPr sz="${sz}"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

export function stepsSlide(): string {
  const title =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>From quote to coverage in four steps</a:t></a:r></a:p></p:txBody></p:sp>';
  const track =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Track"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(128, 408, 1024, 1)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    '<a:solidFill><a:srgbClr val="CFC6B4"/></a:solidFill></p:spPr></p:sp>';
  const steps = LABELS.map((label, j) => {
    const base = 10 + j * 10;
    const x = 64 + j * STEP;
    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${base}" name="Stop ${j + 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrm(x + 107, 387, 43, 43)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
      '<a:solidFill><a:srgbClr val="A87838"/></a:solidFill></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
      textShape(base + 1, `Number ${j + 1}`, [x + 107, 397, 43, 23], String(j + 1), 1000) +
      textShape(base + 2, `Label ${j + 1}`, [x, 467, 256, 45], label, 2000) +
      textShape(base + 3, `Detail ${j + 1}`, [x, 533, 256, 45], `Detail for ${label}.`, 1100)
    );
  }).join('');
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    title +
    track +
    steps +
    '</p:spTree></p:cSld></p:sld>'
  );
}

export interface PictureFixture {
  id: number;
  name: string;
  box: [number, number, number, number];
  /** Relationship the raster blip embeds; `rId3` and `rId4` resolve in the sample deck. */
  embed?: string;
  /** Relationship of a blip that points at its image (`r:link`) instead of embedding it. */
  link?: string;
  /** `a:srcRect` attributes, when the template crops the image to its frame. */
  crop?: string;
  /** Relationship of an `asvg:svgBlip` alternative, which viewers prefer over the raster one. */
  svgEmbed?: string;
  /**
   * The `EG_Media` child making the picture the poster of a clip — `videoFile`,
   * `quickTimeFile`, `audioFile`, `wavAudioFile` or `audioCd`.
   */
  poster?: { tag: string; rel: string };
}

const SVG_BLIP_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_BLIP_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

export function picture({
  id,
  name,
  box,
  embed,
  link,
  crop,
  svgEmbed,
  poster,
}: PictureFixture): string {
  const extLst = svgEmbed
    ? `<a:extLst><a:ext uri="${SVG_BLIP_URI}">` +
      `<asvg:svgBlip xmlns:asvg="${SVG_BLIP_NS}" r:embed="${svgEmbed}"/>` +
      '</a:ext></a:extLst>'
    : '';
  const srcRect = crop ? `<a:srcRect ${crop}/>` : '';
  const raster = embed ? ` r:embed="${embed}"` : '';
  const linkAttr = link ? ` r:link="${link}"` : '';
  const media = poster ? `<a:${poster.tag} r:link="${poster.rel}"/>` : '';
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr>${media}</p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip${raster}${linkAttr}>${extLst}</a:blip>${srcRect}` +
    '<a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}

/** A group holding `child`, occupying `box` on the page. */
export function groupBox(
  id: number,
  name: string,
  box: [number, number, number, number],
  child: string,
): string {
  const at = (value: number) => Math.round(value * EMU);
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${at(box[0])}" y="${at(box[1])}"/>` +
    `<a:ext cx="${at(box[2])}" cy="${at(box[3])}"/>` +
    `<a:chOff x="${at(box[0])}" y="${at(box[1])}"/>` +
    `<a:chExt cx="${at(box[2])}" cy="${at(box[3])}"/></a:xfrm></p:grpSpPr>` +
    child +
    '</p:grpSp>'
  );
}

/** A slide holding a title and the picture `pic` describes, and nothing else. */
export function pictureSlide(pic: string): string {
  return slideXml(textShape(2, 'Title', [64, 64, 800, 60], 'A slide with artwork', 2400) + pic);
}

/** The sample deck with `pic` as the only content of its first slide. */
export async function templateWithPicture(pic: string): Promise<Uint8Array> {
  return templateWithSlide(pictureSlide(pic));
}

export const GROUP: GroupSpec = {
  direction: 'x',
  track: ['3'],
  members: LABELS.map((_, j) => {
    const base = 10 + j * 10;
    return { stop: `${base}`, number: `${base + 1}`, label: `${base + 2}`, detail: `${base + 3}` };
  }),
};

/** The sample deck with its first slide replaced by `slideXml`. */
export async function templateWithSlide(slideXml: string): Promise<Uint8Array> {
  const buf = readFileSync(SRC);
  const bytes = new Uint8Array(buf.length);
  bytes.set(buf);
  const zip = await JSZip.loadAsync(bytes);
  zip.file('ppt/slides/slide1.xml', slideXml, { createFolders: false });
  return zip.generateAsync({ type: 'uint8array' });
}

/** Wrap shape-tree children into a complete slide part. */
export function slideXml(body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>'
  );
}

export async function template(): Promise<Uint8Array> {
  return templateWithSlide(stepsSlide());
}

export async function open(bytes: Uint8Array): Promise<PresentationData> {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  const pres = buildPresentation(await parseZip(copy.buffer));
  pres.slides.forEach((slide) => materializeSlideNodes(pres, slide));
  return pres;
}
