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
