import { buildPresentation } from '../../src/model/Presentation';
import { readSlideElements } from '../../src/editor/readSlideElements';
import { sourceElement } from '../../src/editor/parts';
import { createEditableTextFixture } from './editable-text';

export interface GroupTransform {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export function createGroupedTextFixture(groups: GroupTransform[] = [{}], textRotation = 0) {
  const fixture = createEditableTextFixture();
  const path = fixture.slide.slidePath;
  const text = sourceElement(fixture.slide, '2')!;
  if (textRotation)
    text.getElementsByTagName('a:xfrm')[0].setAttribute('rot', String(textRotation * 60000));
  const source = new XMLSerializer().serializeToString(text);
  let grouped = source;
  for (let index = groups.length - 1; index >= 0; index--) {
    const group = groups[index];
    const id = 10 + index;
    grouped = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="${(group.rotation ?? 0) * 60000}" flipH="${group.flipH ? 1 : 0}" flipV="${group.flipV ? 1 : 0}"><a:off x="${(group.x ?? 200) * 9525}" y="${(group.y ?? 150) * 9525}"/><a:ext cx="${(group.width ?? 600) * 9525}" cy="${(group.height ?? 400) * 9525}"/><a:chOff x="952500" y="952500"/><a:chExt cx="5715000" cy="3810000"/></a:xfrm><a:effectLst><a:glow rad="123"/></a:effectLst></p:grpSpPr>${grouped}</p:grpSp>`;
  }
  const root = fixture.slide.root!.element!;
  const shape = sourceElement(fixture.slide, '2')!;
  const marker = root.ownerDocument.createElement('marker');
  shape.replaceWith(marker);
  const xml = new XMLSerializer().serializeToString(root).replace('<marker/>', grouped);
  fixture.files.slides.set(path, xml);
  const presentation = buildPresentation(fixture.files);
  const slide = presentation.slides[0];
  const elements = readSlideElements(presentation, slide);
  const element = elements.find((item) => item.nodeId === '2')!;
  return { files: fixture.files, presentation, slide, elements, element };
}
