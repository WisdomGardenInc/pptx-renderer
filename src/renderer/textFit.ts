/**
 * Read back how a rendered shape's text fits its box.
 *
 * Writers estimate text size before saving (see `fitText`); the browser is where text is
 * actually laid out. Hosts that generate slides render them, collect each node's element
 * through `onNodeRendered`, and ask here whether the renderer had to shrink the text and
 * whether any of it still spills outside the box — the signal to rewrite or refit.
 *
 * Overflow is judged from the laid-out text itself (a DOM range over the text body)
 * against the text container's box, not from scroll sizes: bottom-anchored text overflows
 * upwards, which scroll sizes never report.
 */

export interface TextFitReport {
  /** Extra shrink the renderer applied to fit the text (1 when none). */
  scale: number;
  /** Text extends past the left or right edge of its box. */
  overflowsX: boolean;
  /** Text extends past the top or bottom edge of its box. */
  overflowsY: boolean;
}

const EDGE_TOLERANCE_PX = 1;

/** Fit report for a rendered node element, or `null` when the node has no text body. */
export function readTextFit(nodeElement: HTMLElement): TextFitReport | null {
  const text = nodeElement.querySelector<HTMLElement>(':scope > [data-pptx-text]');
  if (!text || !text.textContent?.trim()) return null;

  const range = text.ownerDocument.createRange();
  range.selectNodeContents(text);
  const content = range.getBoundingClientRect();
  const box = text.getBoundingClientRect();
  return {
    scale: Number(text.dataset.pptxFitScale ?? '1'),
    overflowsX:
      content.left < box.left - EDGE_TOLERANCE_PX || content.right > box.right + EDGE_TOLERANCE_PX,
    overflowsY:
      content.top < box.top - EDGE_TOLERANCE_PX || content.bottom > box.bottom + EDGE_TOLERANCE_PX,
  };
}
