import { buildPresentation } from '../../src/model/Presentation';
import { parseXml } from '../../src/parser/XmlParser';
import { readSlideElements } from '../../src/editor/readSlideElements';
import { createFillAlphaFixture } from './fill-alpha';

export type ColorMapCase = 'dark' | 'identity' | 'layout' | 'inherit-layout' | 'slide';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const dark = 'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"';
const identity = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"';
const accents = Array.from(
  { length: 6 },
  (_, index) => `accent${index + 1}="accent${index + 1}"`,
).join(' ');

/** Identical scheme slots used by fills, outlines, text and bullets. */
export function createColorMapFixture(mapping: ColorMapCase = 'dark') {
  const { files, slide } = createFillAlphaFixture();
  const shapes = ['bg1', 'tx2'].map((slot, index) => {
    const textSlot = index === 0 ? 'tx2' : 'bg1';
    return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${slot}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${(40 + index * 200) * 9525}" y="381000"/><a:ext cx="1428750" cy="1428750"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="${slot}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="${slot}"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buClr><a:schemeClr val="${textSlot}"/></a:buClr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"><a:solidFill><a:schemeClr val="${textSlot}"/></a:solidFill></a:rPr><a:t>${slot}</a:t></a:r></a:p></p:txBody></p:sp>`;
  });
  const slideOverride =
    mapping === 'slide'
      ? '<p:clrMapOvr><a:overrideClrMapping bg1="dk2" tx2="lt1"/></p:clrMapOvr>'
      : mapping === 'inherit-layout'
        ? '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        : '';
  files.slides.set(
    slide.slidePath,
    `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes.join('')}</p:spTree></p:cSld>${slideOverride}</p:sld>`,
  );
  files.slideMasters.set(
    'ppt/slideMasters/master.xml',
    `<p:sldMaster xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld><p:clrMap ${mapping === 'identity' ? identity : dark} ${accents} hlink="hlink" folHlink="folHlink"/></p:sldMaster>`,
  );
  const layoutOverride = ['layout', 'inherit-layout', 'slide'].includes(mapping)
    ? `<p:clrMapOvr><a:overrideClrMapping ${identity}/></p:clrMapOvr>`
    : '';
  files.slideLayouts.set(
    'ppt/slideLayouts/layout.xml',
    `<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree/></p:cSld>${layoutOverride}</p:sldLayout>`,
  );
  const themePath = 'ppt/theme/theme.xml';
  const theme = parseXml(files.themes.get(themePath)!);
  const scheme = theme.child('themeElements').child('clrScheme').element!;
  for (const [slot, color] of [
    ['dk1', '000000'],
    ['lt1', 'FFFFFF'],
    ['dk2', '1F497D'],
    ['lt2', 'EEECE1'],
  ]) {
    const node = scheme.ownerDocument.createElementNS(A, `a:${slot}`);
    const rgb = scheme.ownerDocument.createElementNS(A, 'a:srgbClr');
    rgb.setAttribute('val', color);
    node.appendChild(rgb);
    scheme.appendChild(node);
  }
  files.themes.set(themePath, new XMLSerializer().serializeToString(theme.element!));
  const presentation = buildPresentation(files);
  return {
    files,
    presentation,
    slide: presentation.slides[0],
    elements: readSlideElements(presentation, presentation.slides[0]),
  };
}
