import { buildPresentation } from '../../src/model/Presentation';
import { createEditableTextFixture } from './editable-text';
import { readSlideElements } from '../../src/editor/readSlideElements';

/** The four alpha-only regressions plus explicit no-fill and pattern controls. */
export function createFillAlphaFixture() {
  const { files, slide } = createEditableTextFixture();
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const fills = [
    '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>',
    '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>',
    '<a:solidFill><a:srgbClr val="FF0000"><a:lumMod val="50000"/></a:srgbClr></a:solidFill>',
    '',
    '<a:noFill/>',
    '<a:pattFill prst="horz"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="0000FF"/></a:bgClr></a:pattFill>',
  ];
  const shapes = fills.map(
    (fill, index) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${index * 200 * 9525}" y="0"/><a:ext cx="1428750" cy="1428750"/></a:xfrm><a:prstGeom prst="rect"/>${fill}<a:ln><a:noFill/></a:ln></p:spPr>${index === 3 ? '<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style>' : ''}</p:sp>`,
  );
  files.slides.set(
    slide.slidePath,
    `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes.join('')}</p:spTree></p:cSld></p:sld>`,
  );
  files.slideRels.set(
    'ppt/slides/_rels/slide1.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/layout.xml"/></Relationships>',
  );
  files.slideLayouts.set(
    'ppt/slideLayouts/layout.xml',
    `<p:sldLayout xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
  );
  files.slideLayoutRels.set(
    'ppt/slideLayouts/_rels/layout.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/master.xml"/></Relationships>',
  );
  files.slideMasters.set(
    'ppt/slideMasters/master.xml',
    `<p:sldMaster xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld><p:clrMap accent1="accent1"/></p:sldMaster>`,
  );
  files.slideMasterRels.set(
    'ppt/slideMasters/_rels/master.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme.xml"/></Relationships>',
  );
  files.themes.set(
    'ppt/theme/theme.xml',
    `<a:theme xmlns:a="${A}" name="Alpha regression"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="4F81BD"/></a:accent1></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`,
  );
  const updated = buildPresentation(files);
  return {
    files,
    presentation: updated,
    slide: updated.slides[0],
    elements: readSlideElements(updated, updated.slides[0]),
  };
}
