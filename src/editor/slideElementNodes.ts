import type { PresentationData } from '../model/Presentation';
import { resolveNodePlaceholderInheritance } from '../model/Presentation';
import { parseRenderableChild } from '../model/RenderableChild';
import type { SlideData, SlideNode } from '../model/Slide';
import type { GroupNodeData } from '../model/nodes/GroupNode';
import {
  remapGroupChild,
  rotationSwapsAxes,
  shouldPropagateGroupFlip,
} from '../model/groupGeometry';
import type { EditableElement } from './EditableElement';

export interface SlideElementNode {
  node: SlideNode;
  parent?: GroupNodeData;
}

function rotate(x: number, y: number, degrees: number): { x: number; y: number } {
  const radians = (degrees * Math.PI) / 180;
  return {
    x: x * Math.cos(radians) - y * Math.sin(radians),
    y: x * Math.sin(radians) + y * Math.cos(radians),
  };
}

export function slideElementNodes(
  slide: SlideData,
  presentation?: PresentationData,
): SlideElementNode[] {
  const layoutPath = presentation?.slideToLayout.get(slide.index);
  const layout = layoutPath ? presentation?.layouts.get(layoutPath) : undefined;
  const masterPath = layoutPath ? presentation?.layoutToMaster.get(layoutPath) : undefined;
  const master = masterPath ? presentation?.masters.get(masterPath) : undefined;
  const result: SlideElementNode[] = [];
  const visit = (original: SlideNode, parent?: GroupNodeData): void => {
    const node = { ...original, position: { ...original.position }, size: { ...original.size } };
    if (parent) {
      remapGroupChild(node, parent);
      const center = rotate(
        node.position.x + node.size.w / 2 - parent.size.w / 2,
        node.position.y + node.size.h / 2 - parent.size.h / 2,
        parent.rotation,
      );
      node.position = {
        x: parent.position.x + parent.size.w / 2 + center.x - node.size.w / 2,
        y: parent.position.y + parent.size.h / 2 + center.y - node.size.h / 2,
      };
      node.rotation += parent.rotation;
    }
    result.push({ node, parent });
    if (node.nodeType !== 'group' || node.source.localName !== 'grpSp') return;
    for (const xml of node.children) {
      const child = parseRenderableChild(xml, {
        rels: slide.rels,
        partPath: slide.slidePath,
        diagramDrawings: presentation?.diagramDrawings,
      });
      if (!child) continue;
      if (presentation)
        resolveNodePlaceholderInheritance(child, layout, master, {
          parentGroup: original as GroupNodeData,
        });
      visit(child, node);
    }
  };
  for (const node of slide.nodes) visit(node);
  return result;
}

export function groupLocalElement(
  element: EditableElement,
  node: SlideNode,
  parent: GroupNodeData,
): EditableElement {
  const rotation = (element.rotation ?? 0) - parent.rotation;
  const center = rotate(
    element.x + element.width / 2 - parent.position.x - parent.size.w / 2,
    element.y + element.height / 2 - parent.position.y - parent.size.h / 2,
    -parent.rotation,
  );
  let x = parent.size.w / 2 + center.x - element.width / 2;
  let y = parent.size.h / 2 + center.y - element.height / 2;
  if (parent.flipH) x = parent.size.w - x - element.width;
  if (parent.flipV) y = parent.size.h - y - element.height;
  const scaleX = parent.childExtent.w > 0 ? parent.size.w / parent.childExtent.w : 1;
  const scaleY = parent.childExtent.h > 0 ? parent.size.h / parent.childExtent.h : 1;
  if (!scaleX || !scaleY) throw new Error('Cannot edit geometry inside a zero-size PPTX group.');
  const offset =
    parent.childExtent.w > 0 || parent.childExtent.h > 0 ? parent.childOffset : { x: 0, y: 0 };
  const swapsAxes = rotationSwapsAxes(rotation);
  const width = element.width / (swapsAxes ? scaleY : scaleX);
  const height = element.height / (swapsAxes ? scaleX : scaleY);
  if (swapsAxes) {
    x = (x + (element.width - element.height) / 2) / scaleX + offset.x - (width - height) / 2;
    y = (y + (element.height - element.width) / 2) / scaleY + offset.y - (height - width) / 2;
  } else {
    x = x / scaleX + offset.x;
    y = y / scaleY + offset.y;
  }
  const flips = shouldPropagateGroupFlip(node);
  return {
    ...element,
    x,
    y,
    width,
    height,
    rotation,
    flipH: flips && parent.flipH ? !element.flipH : element.flipH,
    flipV: flips && parent.flipV ? !element.flipV : element.flipV,
  };
}
