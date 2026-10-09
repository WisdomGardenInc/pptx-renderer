import { buildPresentation } from '../../src/model/Presentation';
import type { PptxFiles } from '../../src/parser/ZipParser';
import { readSlideElements } from '../../src/editor/readSlideElements';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function createEditableTextFixture(
  body = '<a:bodyPr><a:noAutofit/></a:bodyPr>',
  pPr = '',
  content = '',
  options: { empty?: boolean; textBox?: boolean } = {},
) {
  const runs = options.empty
    ? ''
    : '<a:r><a:rPr sz="2400"/><a:t>Hello  </a:t></a:r><a:r><a:rPr sz="2400" b="1"/><a:t>world</a:t></a:r>';
  const shape = (id: number) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr${options.textBox === false ? '' : ' txBox="1"'}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${id === 2 ? 952500 : 7620000}" y="952500"/><a:ext cx="5715000" cy="2857500"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr><p:txBody>${body}<a:lstStyle/><a:p><a:pPr>${pPr}</a:pPr>${runs}${content}</a:p></p:txBody></p:sp>`;
  const files: PptxFiles = {
    contentTypes: '',
    presentation: `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="15240000" cy="8572500"/></p:presentation>`,
    presentationRels: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="${R}/slide" Target="slides/slide1.xml"/></Relationships>`,
    slides: new Map([
      [
        'ppt/slides/slide1.xml',
        `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="bg"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape(2)}${shape(3)}</p:spTree></p:cSld></p:sld>`,
      ],
    ]),
    slideRels: new Map([
      [
        'ppt/slides/_rels/slide1.xml.rels',
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="bg" Type="${R}/image" Target="https://example.com/background.png" TargetMode="External"/></Relationships>`,
      ],
    ]),
    slideLayouts: new Map(),
    slideLayoutRels: new Map(),
    slideMasters: new Map(),
    slideMasterRels: new Map(),
    themes: new Map(),
    media: new Map(),
    charts: new Map(),
    chartStyles: new Map(),
    chartColors: new Map(),
    diagramDrawings: new Map(),
  };
  const presentation = buildPresentation(files);
  const slide = presentation.slides[0];
  return { presentation, slide, element: readSlideElements(presentation, slide)[0] };
}
