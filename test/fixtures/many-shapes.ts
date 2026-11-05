import { buildPresentation } from '../../src/model/Presentation';
import { readSlideElements } from '../../src/editor/readSlideElements';
import { createEditableTextFixture } from './editable-text';

/** One text box following simple solid rectangles, so lookup cannot benefit from its order. */
export function createManyShapesFixture(count: number) {
  const { files, slide } = createEditableTextFixture();
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const rectangles = Array.from(
    { length: count - 1 },
    (_, index) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${index + 3}" name="Rectangle ${index}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${(index % 20) * 381000}" y="${Math.floor(index / 20) * 381000}"/><a:ext cx="285750" cy="285750"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></p:spPr></p:sp>`,
  );
  const target = slide.root!.child('cSld').child('spTree').child('sp').element!;
  files.slides.set(
    slide.slidePath,
    `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${rectangles.join('')}${new XMLSerializer().serializeToString(target)}</p:spTree></p:cSld></p:sld>`,
  );
  files.slideRels.clear();
  const presentation = buildPresentation(files);
  const current = presentation.slides[0];
  const elements = readSlideElements(presentation, current);
  return {
    files,
    presentation,
    slide: current,
    elements,
    element: elements.find((element) => element.nodeId === '2')!,
  };
}
