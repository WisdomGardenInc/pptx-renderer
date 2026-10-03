/**
 * Find repeated groups on a slide — rows of cards, steps of a process, columns of KPIs.
 *
 * A repeated group is several elements that look alike and sit in a regular rhythm. The
 * detection works from that alone, with no knowledge of what the slide says:
 *
 *   1. Style fingerprint: an element's markup with position, size, text, ids and language
 *      tags removed. Two elements with the same fingerprint are drawn the same way.
 *   2. Series: elements sharing a fingerprint, aligned on one edge and evenly spaced along
 *      the other axis.
 *   3. Groups: series with the same axis, count and spacing whose items line up — the k-th
 *      element of each sits at the same offset from the k-th of the others — form one group;
 *      member k takes the k-th element of every series.
 *   4. Track: a textless, thin element running from the first member to the last (the
 *      connector of a timeline) belongs to the group and stretches with it.
 *
 * A single series counts as a group only from three members up, so two lookalikes that
 * happen to line up do not qualify. A group whose series are three or more of the same
 * fingerprint is a grid of cells — a table drawn with shapes — and is left out, as is a
 * group with no text anywhere in it (background stripes, picture frames): there is nothing
 * in it to fill.
 */

import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import { readPlainText } from '../model/nodes/textEdit';
import type { SlideNode } from '../model/Slide';
import type { GroupSpec } from './types';

export interface DetectedGroup extends GroupSpec {
  /** Distance between neighbouring members along the axis, in pixels. */
  spacing: number;
  /** How many members fit at the original spacing before the page margin. */
  maxMembers: number;
}

interface Series {
  fingerprint: string;
  axis: 'x' | 'y';
  nodes: SlideNode[];
  spacing: number;
}

const ALIGN_TOLERANCE_PX = 4;
const SPACING_TOLERANCE_SHARE = 0.06;
const OFFSET_TOLERANCE_PX = 6;
const MIN_LONE_SERIES = 3;
const GRID_SERIES = 3;
const TRACK_THICKNESS_PX = 4;
const PAGE_MARGIN = 0.05;
const POSITION_TAGS = new Set(['off', 'ext', 'chOff', 'chExt']);
const VOLATILE_ATTRS = ['id', 'name', 'lang', 'altLang', 'dirty', 'err', 'smtClean'];

/** Markup with position, size, text, ids and language tags removed. */
export function styleFingerprint(node: SlideNode): string {
  const source = node.source.element;
  if (!source) return `${node.nodeType}:no-source`;
  const clone = source.cloneNode(true) as Element;
  const all = [clone, ...Array.from(clone.getElementsByTagName('*'))];
  for (const el of all) {
    for (const attr of VOLATILE_ATTRS) el.removeAttribute(attr);
    if (el.localName === 't') el.textContent = '';
  }
  for (const el of all) {
    if (POSITION_TAGS.has(el.localName)) el.parentNode?.removeChild(el);
  }
  return `${node.nodeType}:${new XMLSerializer().serializeToString(clone)}`;
}

function start(node: SlideNode, axis: 'x' | 'y'): number {
  return axis === 'x' ? node.position.x : node.position.y;
}

function length(node: SlideNode, axis: 'x' | 'y'): number {
  return axis === 'x' ? node.size.w : node.size.h;
}

function centre(node: SlideNode, axis: 'x' | 'y'): number {
  return start(node, axis) + length(node, axis) / 2;
}

function cross(axis: 'x' | 'y'): 'x' | 'y' {
  return axis === 'x' ? 'y' : 'x';
}

/** Split nodes into runs aligned on the cross axis (same top for rows, same left for columns). */
function alignedRuns(nodes: SlideNode[], axis: 'x' | 'y'): SlideNode[][] {
  const runs: SlideNode[][] = [];
  const sorted = [...nodes].sort((a, b) => start(a, cross(axis)) - start(b, cross(axis)));
  for (const node of sorted) {
    const run = runs.find(
      (candidate) =>
        Math.abs(start(candidate[0], cross(axis)) - start(node, cross(axis))) <= ALIGN_TOLERANCE_PX,
    );
    if (run) run.push(node);
    else runs.push([node]);
  }
  return runs.map((run) => run.sort((a, b) => centre(a, axis) - centre(b, axis)));
}

/** Common spacing of a sorted run, or `null` when the gaps are uneven. */
function evenSpacing(run: SlideNode[], axis: 'x' | 'y'): number | null {
  if (run.length < 2) return null;
  const gaps = run.slice(1).map((node, i) => centre(node, axis) - centre(run[i], axis));
  const mean = gaps.reduce((sum, gap) => sum + gap, 0) / gaps.length;
  if (mean <= 0) return null;
  const tolerance = Math.max(ALIGN_TOLERANCE_PX, mean * SPACING_TOLERANCE_SHARE);
  return gaps.every((gap) => Math.abs(gap - mean) <= tolerance) ? mean : null;
}

function findSeries(nodes: SlideNode[]): Series[] {
  const byFingerprint = new Map<string, SlideNode[]>();
  for (const node of nodes) {
    const key = styleFingerprint(node);
    byFingerprint.set(key, [...(byFingerprint.get(key) ?? []), node]);
  }
  const series: Series[] = [];
  for (const [fingerprint, members] of byFingerprint) {
    if (members.length < 2) continue;
    for (const axis of ['x', 'y'] as const) {
      for (const run of alignedRuns(members, axis)) {
        const spacing = evenSpacing(run, axis);
        if (spacing !== null) series.push({ fingerprint, axis, nodes: run, spacing });
      }
    }
  }
  return series;
}

function lineUp(a: Series, b: Series): boolean {
  if (a.axis !== b.axis || a.nodes.length !== b.nodes.length) return false;
  const tolerance = Math.max(ALIGN_TOLERANCE_PX, a.spacing * SPACING_TOLERANCE_SHARE);
  if (Math.abs(a.spacing - b.spacing) > tolerance) return false;
  const offsets = a.nodes.map((node, i) => centre(b.nodes[i], a.axis) - centre(node, a.axis));
  return offsets.every((offset) => Math.abs(offset - offsets[0]) <= OFFSET_TOLERANCE_PX);
}

/** Cluster series that line up; each node joins at most one cluster. */
function clusterSeries(series: Series[]): Series[][] {
  const ordered = [...series].sort((a, b) => b.nodes.length - a.nodes.length);
  const taken = new Set<SlideNode>();
  const clusters: Series[][] = [];
  for (const candidate of ordered) {
    if (candidate.nodes.some((node) => taken.has(node))) continue;
    const cluster = clusters.find((existing) => lineUp(existing[0], candidate));
    if (cluster) cluster.push(candidate);
    else clusters.push([candidate]);
    candidate.nodes.forEach((node) => taken.add(node));
  }
  return clusters;
}

function hasText(node: SlideNode): boolean {
  return node.nodeType === 'shape' && readPlainText(node as ShapeNodeData).trim() !== '';
}

function carriesText(cluster: Series[]): boolean {
  return cluster.some((series) => series.nodes.some(hasText));
}

function isGrid(cluster: Series[]): boolean {
  const counts = new Map<string, number>();
  for (const series of cluster)
    counts.set(series.fingerprint, (counts.get(series.fingerprint) ?? 0) + 1);
  return [...counts.values()].some((count) => count >= GRID_SERIES);
}

function findTrack(cluster: Series[], others: SlideNode[], members: SlideNode[][]): string[] {
  const axis = cluster[0].axis;
  const first = members[0].map((node) => centre(node, axis));
  const last = members[members.length - 1].map((node) => centre(node, axis));
  const from = Math.min(...first);
  const to = Math.max(...last);
  const all = members.flat();
  const crossLow = Math.min(...all.map((node) => start(node, cross(axis))));
  const crossHigh = Math.max(
    ...all.map((node) => start(node, cross(axis)) + length(node, cross(axis))),
  );
  return others
    .filter((node) => {
      const thin = length(node, cross(axis)) <= TRACK_THICKNESS_PX;
      const spans =
        start(node, axis) <= from + ALIGN_TOLERANCE_PX &&
        start(node, axis) + length(node, axis) >= to - ALIGN_TOLERANCE_PX;
      const inside = start(node, cross(axis)) >= crossLow && start(node, cross(axis)) <= crossHigh;
      return thin && spans && inside;
    })
    .map((node) => node.id);
}

function toGroup(cluster: Series[], others: SlideNode[], pageExtent: number): DetectedGroup {
  const axis = cluster[0].axis;
  const roles = [...cluster].sort(
    (a, b) =>
      start(a.nodes[0], cross(axis)) - start(b.nodes[0], cross(axis)) ||
      start(a.nodes[0], axis) - start(b.nodes[0], axis),
  );
  const count = roles[0].nodes.length;
  const members = Array.from({ length: count }, (_, k) => roles.map((series) => series.nodes[k]));
  const spacing = roles[0].spacing;
  const firstCentre = centre(roles[0].nodes[0], axis);
  const extent = Math.max(...members[0].map((node) => length(node, axis)));
  const room = pageExtent * (1 - PAGE_MARGIN) - extent / 2 - firstCentre;
  return {
    direction: axis,
    spacing,
    maxMembers: Math.max(count, Math.floor(room / spacing) + 1),
    members: members.map((nodes) =>
      Object.fromEntries(nodes.map((node, i) => [`e${i + 1}`, node.id])),
    ),
    track: findTrack(cluster, others, members),
  };
}

/** Repeated groups among a slide's top-level nodes. */
export function detectGroups(
  nodes: readonly SlideNode[],
  page: { width: number; height: number },
): DetectedGroup[] {
  const clusters = clusterSeries(findSeries([...nodes])).filter(
    (cluster) =>
      (cluster.length >= 2 || cluster[0].nodes.length >= MIN_LONE_SERIES) &&
      !isGrid(cluster) &&
      carriesText(cluster),
  );
  const grouped = new Set(clusters.flatMap((cluster) => cluster.flatMap((series) => series.nodes)));
  const others = nodes.filter((node) => !grouped.has(node));
  return clusters.map((cluster) =>
    toGroup(cluster, others, cluster[0].axis === 'x' ? page.width : page.height),
  );
}
