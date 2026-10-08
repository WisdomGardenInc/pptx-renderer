/**
 * Every element of one template slide, group children included.
 *
 * A group's children are raw XML until something parses them, so a reader that stops at
 * `slide.nodes` sees only the group itself — the cards, agenda entries and header tabs
 * inside it stay invisible. Both the page analysis and the edit executor need the same
 * view: each element's id, the groups it sits in, and where it lands on the page. Parsing
 * it once here keeps the reported ids and the executed ids identical.
 */

import type { PresentationData } from '../model/Presentation';
import { resolveNodePlaceholderInheritance } from '../model/Presentation';
import type { SlideData, SlideNode } from '../model/Slide';
import {
  groupChildTransform,
  IDENTITY_TRANSFORM,
  transformedBounds,
  type Bounds,
  type CoordinateTransform,
} from '../model/groupTransform';
import type { BaseNodeData } from '../model/nodes/BaseNode';
import type { GroupNodeData } from '../model/nodes/GroupNode';
import { parseRenderableChild } from '../model/RenderableChild';

export interface TemplateElement {
  node: SlideNode | BaseNodeData;
  /** Ids of the groups the element sits in, outermost first; empty when top-level. */
  parentIds: string[];
  /** Where the element lands on the page, after every enclosing group transform. */
  bounds: Bounds;
  /** Maps the element's own coordinates onto the page. */
  transform: CoordinateTransform;
}

export interface TemplateElements {
  all: TemplateElement[];
  byId: Map<string, TemplateElement>;
}

function parseContext(presentation: PresentationData, slide: SlideData) {
  const layoutPath = presentation.slideToLayout.get(slide.index) || slide.layoutIndex;
  const layout = presentation.layouts.get(layoutPath);
  const masterPath = layoutPath ? presentation.layoutToMaster.get(layoutPath) : '';
  const master = masterPath ? presentation.masters.get(masterPath) : undefined;
  return {
    rels: slide.rels,
    partPath: slide.slidePath,
    diagramDrawings: presentation.diagramDrawings,
    layout,
    master,
  };
}

export function templateElements(
  presentation: PresentationData,
  slide: SlideData,
): TemplateElements {
  const ctx = parseContext(presentation, slide);
  const all: TemplateElement[] = [];

  const visit = (
    node: SlideNode | BaseNodeData,
    parentIds: string[],
    transform: CoordinateTransform,
  ): void => {
    all.push({ node, parentIds, bounds: transformedBounds(node, transform), transform });
    if (node.nodeType !== 'group') return;
    const group = node as GroupNodeData;
    for (const childXml of group.children) {
      let child: BaseNodeData | undefined;
      try {
        child = parseRenderableChild(childXml, { ...ctx, skipPlaceholders: false });
      } catch {
        child = undefined;
      }
      if (!child) continue;
      resolveNodePlaceholderInheritance(child, ctx.layout, ctx.master, { parentGroup: group });
      visit(child, [...parentIds, group.id], groupChildTransform(group, child, transform));
    }
  };

  for (const node of slide.nodes) visit(node, [], IDENTITY_TRANSFORM);

  return { all, byId: new Map(all.map((entry) => [entry.node.id, entry])) };
}
