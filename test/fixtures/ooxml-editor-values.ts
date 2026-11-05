import { buildPresentation } from '../../src/model/Presentation';
import { parseXml } from '../../src/parser/XmlParser';
import { readSlideElements } from '../../src/editor/readSlideElements';
import { createEditableTextFixture } from './editable-text';
import { createFillAlphaFixture } from './fill-alpha';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

export function createOoxmlEditorFixture(
  value = 'true',
  spacing = '<a:spcBef><a:spcPct val="200000"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft>',
) {
  const { files, slide } = createEditableTextFixture();
  const root = parseXml(files.slides.get(slide.slidePath)!);
  root.child('cSld').child('bg').element!.remove();
  for (const shape of root.child('cSld').child('spTree').children('sp')) {
    const cNv = shape.child('nvSpPr').child('cNvSpPr').element!;
    cNv.setAttribute('txBox', value);
    const locks = cNv.ownerDocument.createElementNS(A, 'a:spLocks');
    locks.setAttribute('noMove', value);
    locks.setAttribute('noResize', value);
    cNv.appendChild(locks);
    const body = shape.child('txBody').element!;
    shape.child('txBody').child('p').element!.remove();
    for (const text of ['First', 'Middle', 'Last']) {
      const paragraph = parseXml(
        `<a:p xmlns:a="${A}"><a:pPr>${spacing}<a:defRPr b="true" i="true"/></a:pPr><a:r><a:rPr sz="2400" b="${value}" i="${value}"/><a:t>${text}</a:t></a:r></a:p>`,
      );
      body.appendChild(body.ownerDocument.importNode(paragraph.element!, true));
    }
  }
  files.slides.set(slide.slidePath, new XMLSerializer().serializeToString(root.element!));
  files.slideRels.clear();
  const presentation = buildPresentation(files);
  return {
    files,
    presentation,
    slide: presentation.slides[0],
    elements: readSlideElements(presentation, presentation.slides[0]),
  };
}

export function createPlaceholderStyleFixture(defaultIndex = false) {
  const { files, slide } = createFillAlphaFixture();
  const ph = defaultIndex ? '' : ' idx="7"';
  const shape = (id: number, type: string, index: string, style: string, ownText = '') =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}"/><p:cNvSpPr/><p:nvPr><p:ph${type ? ` type="${type}"` : ''}${index}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="3810000"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle>${style}</a:lstStyle><a:p>${ownText}</a:p></p:txBody></p:sp>`;
  // Local title does not rewire the matched layout's body-category master parent.
  files.slides.set(
    slide.slidePath,
    `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape(2, 'title', ph, '', '<a:r><a:rPr/><a:t>Inherited</a:t></a:r>')}</p:spTree></p:cSld></p:sld>`,
  );
  files.slideLayouts.set(
    'ppt/slideLayouts/layout.xml',
    `<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape(10, defaultIndex ? '' : 'obj', ph, '<a:lvl1pPr algn="r"><a:spcBef><a:spcPct val="150000"/></a:spcBef><a:defRPr i="true"/></a:lvl1pPr>')}</p:spTree></p:cSld></p:sldLayout>`,
  );
  files.slideMasters.set(
    'ppt/slideMasters/master.xml',
    `<p:sldMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape(20, 'title', ph, '<a:lvl1pPr><a:defRPr b="false" sz="4400"/></a:lvl1pPr>')}${shape(21, 'body', ' idx="99"', '<a:lvl1pPr><a:defRPr b="true" sz="2800"/></a:lvl1pPr>')}</p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4000"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:spcAft><a:spcPts val="600"/></a:spcAft></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>`,
  );
  const presentation = buildPresentation(files);
  return {
    files,
    presentation,
    slide: presentation.slides[0],
    elements: readSlideElements(presentation, presentation.slides[0]),
  };
}
