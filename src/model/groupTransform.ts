import type { BaseNodeData, Position, Size } from './nodes/BaseNode';
import type { GroupNodeData } from './nodes/GroupNode';

export interface CoordinateTransform {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
}

export interface Bounds extends Position, Size {}

export const IDENTITY_TRANSFORM: CoordinateTransform = {
  offsetX: 0,
  offsetY: 0,
  scaleX: 1,
  scaleY: 1,
};

const rotationSwapsAxes = (rotation: number): boolean => {
  const normalized = ((rotation % 360) + 360) % 360;
  return Math.abs(normalized - 90) < 0.0001 || Math.abs(normalized - 270) < 0.0001;
};

export function transformedBounds(
  node: BaseNodeData,
  transform: CoordinateTransform = IDENTITY_TRANSFORM,
): Bounds {
  return {
    x: transform.offsetX + node.position.x * transform.scaleX,
    y: transform.offsetY + node.position.y * transform.scaleY,
    w: node.size.w * transform.scaleX,
    h: node.size.h * transform.scaleY,
  };
}

export function groupChildTransform(
  group: GroupNodeData,
  child: BaseNodeData,
  parentTransform: CoordinateTransform,
): CoordinateTransform {
  if (group.childExtent.w <= 0 && group.childExtent.h <= 0) {
    return {
      offsetX: parentTransform.offsetX + group.position.x * parentTransform.scaleX,
      offsetY: parentTransform.offsetY + group.position.y * parentTransform.scaleY,
      scaleX: parentTransform.scaleX,
      scaleY: parentTransform.scaleY,
    };
  }

  const scaleX = group.childExtent.w > 0 ? group.size.w / group.childExtent.w : 1;
  const scaleY = group.childExtent.h > 0 ? group.size.h / group.childExtent.h : 1;
  if (rotationSwapsAxes(child.rotation)) {
    const rotatedBBoxX = child.position.x + (child.size.w - child.size.h) / 2;
    const rotatedBBoxY = child.position.y + (child.size.h - child.size.w) / 2;
    const nextSize = {
      w: child.size.w * scaleY,
      h: child.size.h * scaleX,
    };
    const nextPosition = {
      x: (rotatedBBoxX - group.childOffset.x) * scaleX - (nextSize.w - nextSize.h) / 2,
      y: (rotatedBBoxY - group.childOffset.y) * scaleY - (nextSize.h - nextSize.w) / 2,
    };
    const childScaleX = parentTransform.scaleX * scaleY;
    const childScaleY = parentTransform.scaleY * scaleX;
    return {
      offsetX:
        parentTransform.offsetX +
        (group.position.x + nextPosition.x) * parentTransform.scaleX -
        child.position.x * childScaleX,
      offsetY:
        parentTransform.offsetY +
        (group.position.y + nextPosition.y) * parentTransform.scaleY -
        child.position.y * childScaleY,
      scaleX: childScaleX,
      scaleY: childScaleY,
    };
  }

  return {
    offsetX:
      parentTransform.offsetX +
      (group.position.x - group.childOffset.x * scaleX) * parentTransform.scaleX,
    offsetY:
      parentTransform.offsetY +
      (group.position.y - group.childOffset.y * scaleY) * parentTransform.scaleY,
    scaleX: parentTransform.scaleX * scaleX,
    scaleY: parentTransform.scaleY * scaleY,
  };
}
