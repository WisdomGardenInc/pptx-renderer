import type { TextFitAction } from '../writer/fitText';

/**
 * A repeated group on a template slide — cards, steps, KPIs — described by the caller.
 * Each member maps a role (`number`, `label`, `detail`, …) to the id of the top-level
 * element playing it; every member lists the same roles.
 */
export interface GroupSpec {
  members: Array<Record<string, string>>;
  /** Axis the members are laid out along. */
  direction: 'x' | 'y';
  /** Elements spanning the whole group (a timeline's connector); resized with it. */
  track?: string[];
}

export type ComposeOp =
  /** Replace an element's text, keeping its styling. */
  | { op: 'set_text'; element: string; text: string }
  /** Empty an element's text. */
  | { op: 'clear'; element: string }
  /** Fill a group with one entry per item, adding or removing members to match. */
  | { op: 'set_items'; group: GroupSpec; items: Array<Record<string, string>> };

export interface PagePlan {
  /** 1-based position of the template slide this page is built from. */
  source: number;
  ops: ComposeOp[];
}

export interface ComposedText {
  /** 0-based page in the composed deck. */
  page: number;
  element: string;
  action: TextFitAction;
  scale: number;
}

export interface ComposeProblem {
  page: number;
  message: string;
}

export interface ComposeResult {
  bytes: Uint8Array;
  texts: ComposedText[];
  problems: ComposeProblem[];
}
