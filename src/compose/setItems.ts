/**
 * `set_items`: make a repeated group hold exactly as many members as there are items.
 *
 * Members past the item count are removed. Missing ones are copies of the last member,
 * stepped on by the group's own spacing. The group then keeps its rhythm: members extend
 * along the axis at the original spacing while the page has room, and the spacing tightens
 * only when it does not — a group that shrinks re-spreads across the span it had, so three
 * steps still fill the width four used to. Track elements (a connector running through
 * every member) stretch or shrink with the distance between the first and last member.
 */

import type { SlideData, SlideNode } from '../model/Slide';
import { duplicateNodes } from '../writer/duplicateNodes';
import type { GroupSpec } from './types';

type Member = Map<string, SlideNode>;

interface ItemsLayout {
  /** Members after the change, in order, each mapping role → node. */
  members: Member[];
  /** Human-readable notes about compromises (tight spacing, missing elements). */
  notes: string[];
}

const PAGE_MARGIN = 0.05;

function centre(nodes: Iterable<SlideNode>, axis: 'x' | 'y'): number {
  let low = Number.POSITIVE_INFINITY;
  let high = Number.NEGATIVE_INFINITY;
  for (const node of nodes) {
    const start = axis === 'x' ? node.position.x : node.position.y;
    const extent = axis === 'x' ? node.size.w : node.size.h;
    low = Math.min(low, start);
    high = Math.max(high, start + extent);
  }
  return (low + high) / 2;
}

function extent(nodes: Iterable<SlideNode>, axis: 'x' | 'y'): number {
  let low = Number.POSITIVE_INFINITY;
  let high = Number.NEGATIVE_INFINITY;
  for (const node of nodes) {
    const start = axis === 'x' ? node.position.x : node.position.y;
    low = Math.min(low, start);
    high = Math.max(high, start + (axis === 'x' ? node.size.w : node.size.h));
  }
  return high - low;
}

function shift(node: SlideNode, axis: 'x' | 'y', by: number): void {
  node.position =
    axis === 'x'
      ? { ...node.position, x: node.position.x + by }
      : { ...node.position, y: node.position.y + by };
}

function resolveMembers(slide: SlideData, group: GroupSpec): Member[] | string {
  const byId = new Map(slide.nodes.map((node) => [node.id, node]));
  const members: Member[] = [];
  for (const spec of group.members) {
    const member: Member = new Map();
    for (const [role, id] of Object.entries(spec)) {
      const node = byId.get(id);
      if (!node) return `group element ${id} (${role}) is not a top-level element of the slide`;
      member.set(role, node);
    }
    members.push(member);
  }
  return members.length ? members : 'group has no members';
}

/** Lay `group` out for `count` items, or return why it cannot be done. */
export function layoutItems(
  slide: SlideData,
  group: GroupSpec,
  count: number,
  page: { width: number; height: number },
): ItemsLayout | string {
  const resolved = resolveMembers(slide, group);
  if (typeof resolved === 'string') return resolved;
  const axis = group.direction;
  const notes: string[] = [];
  const centres = resolved.map((member) => centre(member.values(), axis));
  const first = centres[0];
  const originalLast = centres[centres.length - 1];
  const spacing =
    resolved.length > 1
      ? (originalLast - first) / (resolved.length - 1)
      : extent(resolved[0].values(), axis);

  const dropped = resolved.slice(count).flatMap((member) => [...member.values()]);
  slide.nodes = slide.nodes.filter((node) => !dropped.includes(node));
  const members = resolved.slice(0, Math.max(count, 0));

  while (members.length < count) {
    const template = members[members.length - 1];
    const roles = [...template.keys()];
    const offset = axis === 'x' ? { dx: spacing, dy: 0 } : { dx: 0, dy: spacing };
    const copies = duplicateNodes(
      slide,
      roles.map((role) => template.get(role)!),
      offset,
    );
    members.push(new Map(roles.map((role, index) => [role, copies[index]])));
  }
  if (!members.length) return { members, notes };

  const memberExtent = extent(members[0].values(), axis);
  const pageEnd = (axis === 'x' ? page.width : page.height) * (1 - PAGE_MARGIN);
  let last = originalLast;
  if (members.length > 1 && members.length > resolved.length) {
    const wanted = first + spacing * (members.length - 1);
    last = Math.max(originalLast, Math.min(wanted, pageEnd - memberExtent / 2));
  }
  const step = members.length > 1 ? (last - first) / (members.length - 1) : 0;
  if (members.length > 1 && step < memberExtent) {
    notes.push(`${members.length} items are tighter than one item's width; neighbours may overlap`);
  }

  members.forEach((member, index) => {
    const target = first + step * index;
    const by = target - centre(member.values(), axis);
    for (const node of member.values()) shift(node, axis, by);
  });

  const byId = new Map(slide.nodes.map((node) => [node.id, node]));
  const stretch = (members.length > 1 ? last : first) - originalLast;
  for (const id of group.track ?? []) {
    const track = byId.get(id);
    if (!track) {
      notes.push(`track element ${id} is not on the slide`);
      continue;
    }
    track.size =
      axis === 'x'
        ? { ...track.size, w: Math.max(1, track.size.w + stretch) }
        : { ...track.size, h: Math.max(1, track.size.h + stretch) };
  }
  return { members, notes };
}
