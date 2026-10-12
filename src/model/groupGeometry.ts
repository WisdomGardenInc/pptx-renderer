import type { BaseNodeData } from './nodes/BaseNode';
import type { GroupNodeData } from './nodes/GroupNode';
import type { ShapeNodeData } from './nodes/ShapeNode';

export function shouldPropagateGroupFlip(node: BaseNodeData): boolean {
  return node.nodeType !== 'table' && node.nodeType !== 'chart';
}

export function rotationSwapsAxes(rotation: number): boolean {
  const normalized = ((rotation % 360) + 360) % 360;
  return Math.abs(normalized - 90) < 0.0001 || Math.abs(normalized - 270) < 0.0001;
}

export function remapGroupChild(child: BaseNodeData, group: GroupNodeData): void {
  const scaleX = group.childExtent.w > 0 ? group.size.w / group.childExtent.w : 1;
  const scaleY = group.childExtent.h > 0 ? group.size.h / group.childExtent.h : 1;
  const offset =
    group.childExtent.w > 0 || group.childExtent.h > 0 ? group.childOffset : { x: 0, y: 0 };
  const swapsAxes = rotationSwapsAxes(child.rotation);
  const position = child.position;
  const size = child.size;
  if (swapsAxes) {
    const nextSize = { w: size.w * scaleY, h: size.h * scaleX };
    child.position = {
      x: (position.x + (size.w - size.h) / 2 - offset.x) * scaleX - (nextSize.w - nextSize.h) / 2,
      y: (position.y + (size.h - size.w) / 2 - offset.y) * scaleY - (nextSize.h - nextSize.w) / 2,
    };
    child.size = nextSize;
  } else {
    child.position = {
      x: (position.x - offset.x) * scaleX,
      y: (position.y - offset.y) * scaleY,
    };
    child.size = { w: size.w * scaleX, h: size.h * scaleY };
  }
  if (child.nodeType === 'shape') {
    const shape = child as ShapeNodeData;
    const box = shape.textBoxBounds;
    if (box) {
      const sx = swapsAxes ? scaleY : scaleX;
      const sy = swapsAxes ? scaleX : scaleY;
      shape.textBoxBounds = { x: box.x * sx, y: box.y * sy, w: box.w * sx, h: box.h * sy };
    }
  }
  if (group.flipH) {
    child.position = { ...child.position, x: group.size.w - child.position.x - child.size.w };
    if (shouldPropagateGroupFlip(child)) child.flipH = !child.flipH;
  }
  if (group.flipV) {
    child.position = { ...child.position, y: group.size.h - child.position.y - child.size.h };
    if (shouldPropagateGroupFlip(child)) child.flipV = !child.flipV;
  }
}
