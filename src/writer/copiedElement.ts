import type { EditableElement } from '../editor/EditableElement';
import { R } from './xmlEdit';

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

export function normalizeCopiedRelationships(source: Element): void {
  let hasRelationships = false;
  for (const node of [source, ...Array.from(source.getElementsByTagName('*'))]) {
    for (const attribute of Array.from(node.attributes)) {
      if (attribute.namespaceURI !== R) continue;
      node.removeAttributeNode(attribute);
      node.setAttributeNS(R, `r:${attribute.localName}`, attribute.value);
      hasRelationships = true;
    }
  }
  if (hasRelationships) source.setAttributeNS(XMLNS_NS, 'xmlns:r', R);
}

export function copiedStyleBaseline(element: EditableElement): EditableElement {
  if (element.copyStyle !== 'snapshot') return element;
  return {
    ...element,
    paragraphs: undefined,
    vAlign: undefined,
    autoFit: undefined,
    fontScale: undefined,
    lineSpaceReduction: undefined,
    wrap: undefined,
    bodyPadding: undefined,
    opacity: undefined,
    picCrop: undefined,
    cornerRadius: undefined,
    lineCap: undefined,
    headEnd: undefined,
    tailEnd: undefined,
    ...(element.fillColor ? { fillColor: undefined, fillAlpha: undefined } : {}),
    ...(element.strokeColor
      ? {
          strokeColor: undefined,
          strokeAlpha: undefined,
          strokeWidth: undefined,
          strokeDash: undefined,
        }
      : {}),
  };
}
