/** DOM adapters for text emitted by renderEditableText(). No host model or framework dependency. */
export interface EditableTextRunPosition {
  readonly paraIdx: number;
  readonly runIdx: number;
  readonly charOffset: number;
}

export interface EditableTextParagraphPosition {
  readonly paraIdx: number;
  readonly flatChar: number;
}

export function normalizeEditableWhitespace(
  root: HTMLElement,
  sourceParagraphs: readonly { runs: readonly { text: string }[] }[],
): void {
  const selection = root.ownerDocument.getSelection();
  const anchor = selection?.anchorNode;
  const focus = selection?.focusNode;
  const anchorOffset = selection?.anchorOffset ?? 0;
  const focusOffset = selection?.focusOffset ?? 0;
  const restoreSelection = anchor && focus && root.contains(anchor) && root.contains(focus);
  let changed = false;
  const walker = root.ownerDocument.createTreeWalker(root, 4);
  for (let text = walker.nextNode(); text; text = walker.nextNode()) {
    const parent = text.parentElement;
    if (parent?.closest('[data-bullet="true"]')) continue;
    const paragraph = parent?.closest<HTMLElement>('[data-pptx-paragraph]');
    const run = parent?.closest<HTMLElement>('[data-pptx-run]');
    const source = sourceParagraphs[Number(paragraph?.dataset.pptxParagraph ?? 0)];
    const sourceRun = source?.runs[Number(run?.dataset.pptxRun ?? 0)];
    if (sourceRun?.text.includes('\t')) continue;
    const before = text.nodeValue ?? '';
    const normalized = before.replace(/[ \u00a0]+/g, (whitespace, index: number) => {
      const end = index + whitespace.length;
      // Native editing uses NBSP at text boundaries and around the caret. Leave
      // those spaces intact until typing moves beyond them.
      if (
        index === 0 ||
        end === before.length ||
        (selection?.isCollapsed && anchor === text && index <= anchorOffset && anchorOffset <= end)
      )
        return whitespace;
      const plain = sourceRun?.text.includes('\u00a0')
        ? whitespace
        : whitespace.replace(/\u00a0/g, ' ');
      return plain.replace(/ {2}/g, ' \u00a0');
    });
    if (normalized === before) continue;
    // Replace only the changed characters; assigning nodeValue resets native edit offsets.
    const data = text as Text;
    for (let index = 0; index < normalized.length; index++) {
      if (normalized[index] !== before[index]) data.replaceData(index, 1, normalized[index]);
    }
    changed = true;
  }
  if (
    changed &&
    restoreSelection &&
    selection &&
    (selection.anchorNode !== anchor ||
      selection.anchorOffset !== anchorOffset ||
      selection.focusNode !== focus ||
      selection.focusOffset !== focusOffset)
  ) {
    selection.setBaseAndExtent(anchor, anchorOffset, focus, focusOffset);
  }
}

export function isEditableTextRoot(root: HTMLElement): boolean {
  return root.dataset.pptxEditable === 'true';
}

const ZWSP = /\u200B/g;

function stripZwsp(text: string): string {
  return text.replace(ZWSP, '');
}

function getParagraphEls(root: HTMLElement): HTMLElement[] {
  return Array.from(root.children).filter((c): c is HTMLElement =>
    c.hasAttribute('data-pptx-paragraph'),
  );
}

function getRunSpans(paragraph: HTMLElement): HTMLElement[] {
  return Array.from(paragraph.querySelectorAll<HTMLElement>('[data-pptx-run]'));
}

function runText(element: HTMLElement): string {
  return element.tagName === 'BR' ? '\n' : stripZwsp(element.textContent ?? '');
}

function findAncestorTag(node: Node, tag: string, root: HTMLElement): HTMLElement | null {
  let cur: Node | null = node;
  while (cur && cur !== root) {
    if (
      cur.nodeType === Node.ELEMENT_NODE &&
      tag === 'P' &&
      (cur as HTMLElement).hasAttribute('data-pptx-paragraph')
    ) {
      return cur as HTMLElement;
    }
    cur = cur.parentNode;
  }
  return null;
}

function preTextLengthBefore(target: Node, targetOffset: number, container: HTMLElement): number {
  let count = 0;
  let found = false;
  function visitChildrenUpTo(parent: Node, limit: number): void {
    let i = 0;
    for (const c of Array.from(parent.childNodes)) {
      if (i >= limit || found) break;
      visit(c);
      i++;
    }
  }
  function handleTargetNode(n: Node): void {
    if (target.nodeType === Node.TEXT_NODE) {
      count += stripZwsp((target.textContent ?? '').slice(0, targetOffset)).length;
    } else {
      visitChildrenUpTo(n, targetOffset);
    }
    found = true;
  }
  function visit(n: Node): void {
    if (found) return;
    if (n === target) {
      handleTargetNode(n);
      return;
    }
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).getAttribute('data-bullet') === 'true'
    )
      return;
    if (
      n.nodeType === Node.ELEMENT_NODE &&
      (n as HTMLElement).tagName === 'BR' &&
      (n as HTMLElement).hasAttribute('data-pptx-run')
    ) {
      count++;
      return;
    }
    if (n.nodeType === Node.TEXT_NODE) {
      count += stripZwsp(n.textContent ?? '').length;
      return;
    }
    for (const c of Array.from(n.childNodes)) {
      if (found) return;
      visit(c);
    }
  }
  visit(container);
  return count;
}

export function mapEditableTextPositionToRun(
  node: Node,
  offset: number,
  root: HTMLElement,
): EditableTextRunPosition | null {
  const pEl = findAncestorTag(node, 'P', root);
  if (!pEl) return null;
  const pEls = getParagraphEls(root);
  const paraIdx = pEls.indexOf(pEl);
  if (paraIdx < 0) return null;

  const runSpans = getRunSpans(pEl);
  if (runSpans.length === 0) {
    const charOffset = preTextLengthBefore(node, offset, pEl);
    return { paraIdx, runIdx: 0, charOffset };
  }

  let spanEl: HTMLElement | null =
    runSpans.find((span) => span === node || span.contains(node)) ??
    findAncestorTag(node, 'SPAN', pEl);
  if (spanEl?.getAttribute('data-bullet') === 'true') spanEl = null;

  if (spanEl && runSpans.includes(spanEl)) {
    const runIdx = runSpans.indexOf(spanEl);
    const charOffset = preTextLengthBefore(node, offset, spanEl);
    return { paraIdx, runIdx, charOffset };
  }

  const totalBefore = preTextLengthBefore(node, offset, pEl);
  let acc = 0;
  for (let i = 0; i < runSpans.length; i++) {
    const len = runText(runSpans[i]).length;
    if (totalBefore <= acc + len) {
      return { paraIdx, runIdx: i, charOffset: totalBefore - acc };
    }
    acc += len;
  }
  const last = runSpans.length - 1;
  return { paraIdx, runIdx: last, charOffset: runText(runSpans[last]).length };
}

export function mapEditableTextPositionToParagraph(
  node: Node,
  offset: number,
  root: HTMLElement,
): { paraIdx: number; flatChar: number } | null {
  const pEl = findAncestorTag(node, 'P', root);
  if (!pEl) return null;
  const paraIdx = getParagraphEls(root).indexOf(pEl);
  if (paraIdx < 0) return null;
  return { paraIdx, flatChar: preTextLengthBefore(node, offset, pEl) };
}

function findOffsetInTextNode(textNode: Node, localOffset: number): number {
  const raw = textNode.textContent ?? '';
  let i = 0;
  let counted = 0;
  while (counted < localOffset && i < raw.length) {
    if (raw[i] !== '\u200B') counted++;
    i++;
  }
  return i;
}

function positionInsideSpan(
  span: HTMLElement,
  localOffset: number,
): { node: Node; offset: number } {
  if (span.tagName === 'BR' && span.parentNode) {
    const index = Array.from(span.parentNode.childNodes).indexOf(span);
    return { node: span.parentNode, offset: index + (localOffset > 0 ? 1 : 0) };
  }
  const walker = span.ownerDocument.createTreeWalker(span, 4 /* SHOW_TEXT */);
  let remaining = localOffset;
  let last: Node | null = null;
  for (let textNode = walker.nextNode(); textNode; textNode = walker.nextNode()) {
    last = textNode;
    const length = stripZwsp(textNode.textContent ?? '').length;
    if (remaining <= length)
      return {
        node: textNode,
        offset: findOffsetInTextNode(textNode, remaining),
      };
    remaining -= length;
  }
  if (last) return { node: last, offset: (last.textContent ?? '').length };
  return { node: span, offset: 0 };
}

function positionAtEndOfLastSpan(runSpans: HTMLElement[]): {
  node: Node;
  offset: number;
} {
  const lastSpan = runSpans[runSpans.length - 1];
  return positionInsideSpan(lastSpan, runText(lastSpan).length);
}

function findDomPositionForFlatChar(
  root: HTMLElement,
  paraIdx: number,
  flatChar: number,
): { node: Node; offset: number } | null {
  const pEls = getParagraphEls(root);
  const pEl = pEls[paraIdx];
  if (!pEl) return null;
  const runSpans = getRunSpans(pEl);
  if (runSpans.length === 0) return { node: pEl, offset: 0 };
  let acc = 0;
  for (let i = 0; i < runSpans.length; i++) {
    const span = runSpans[i];
    const len = runText(span).length;
    if (flatChar < acc + len) return positionInsideSpan(span, flatChar - acc);
    if (flatChar === acc + len) {
      const isLast = i === runSpans.length - 1;
      return isLast ? positionInsideSpan(span, len) : positionInsideSpan(runSpans[i + 1], 0);
    }
    acc += len;
  }
  return positionAtEndOfLastSpan(runSpans);
}

export function createEditableTextRange(
  root: HTMLElement,
  startParaIdx: number,
  startFlatChar: number,
  endParaIdx: number,
  endFlatChar: number,
): Range | null {
  const start = findDomPositionForFlatChar(root, startParaIdx, startFlatChar);
  const end = findDomPositionForFlatChar(root, endParaIdx, endFlatChar);
  if (!start || !end) return null;
  const range = root.ownerDocument.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  return range;
}

interface EditableWhitespaceSnapshot {
  readonly text: string;
  readonly content: string;
}

function paragraphTextNodes(paragraph: HTMLElement): Node[] {
  const nodes: Node[] = [];
  const walker = paragraph.ownerDocument.createTreeWalker(paragraph, 5, {
    acceptNode(node) {
      if (node.nodeType === 3) return 1;
      const element = node as HTMLElement;
      if (element.dataset.bullet === 'true') return 2;
      return element.tagName === 'BR' && element.hasAttribute('data-pptx-run') ? 1 : 3;
    },
  });
  for (let node = walker.nextNode(); node; node = walker.nextNode()) nodes.push(node);
  return nodes;
}

function paragraphNodeText(nodes: readonly Node[]): string {
  return nodes
    .map((node) => (node.nodeType === 3 ? stripZwsp(node.nodeValue ?? '') : '\n'))
    .join('');
}

export function captureEditableWhitespace(
  root: HTMLElement,
  sourceParagraphs: readonly { runs: { text: string }[] }[],
): EditableWhitespaceSnapshot[] {
  const paragraphs = extractEditableText(root, sourceParagraphs);
  return getParagraphEls(root).map((paragraph, index) => ({
    text: paragraphNodeText(paragraphTextNodes(paragraph)),
    content: paragraphs[index].runs.map((run) => run.text).join(''),
  }));
}

export function restoreEditableWhitespace(
  root: HTMLElement,
  snapshot: readonly EditableWhitespaceSnapshot[],
  paragraphs: readonly { runs: readonly { text: string }[] }[],
): void {
  for (const [index, paragraph] of getParagraphEls(root).entries()) {
    const previous = snapshot[index];
    if (!previous || previous.content !== paragraphs[index]?.runs.map((run) => run.text).join(''))
      continue;
    const nodes = paragraphTextNodes(paragraph);
    const text = paragraphNodeText(nodes);
    if (text.replace(/\u00a0/g, ' ') !== previous.text.replace(/\u00a0/g, ' ')) continue;
    let flatChar = 0;
    for (const node of nodes) {
      if (node.nodeType !== 3) {
        flatChar++;
        continue;
      }
      const before = node.nodeValue ?? '';
      for (let offset = 0; offset < before.length; offset++) {
        if (before[offset] === '\u200b') continue;
        const original = previous.text[flatChar++];
        if (before[offset] !== original && /[ \u00a0]/.test(before[offset]))
          (node as Text).replaceData(offset, 1, original);
      }
    }
  }
}

/** Read text while retaining the host's paragraph/run metadata (including bullet definitions). */
export function extractEditableText<P extends { runs: { text: string }[] }>(
  root: HTMLElement,
  sourceParagraphs: readonly P[],
): P[] {
  return getParagraphEls(root).map((paragraph, paragraphIndex) => {
    const source =
      sourceParagraphs[paragraphIndex] ?? sourceParagraphs[sourceParagraphs.length - 1];
    const sourceRuns = source?.runs ?? [];
    const runs: { text: string }[] = [];
    let index = 0;
    const append = (text: string, sourceIndex: number) => {
      const run = sourceRuns[sourceIndex] ?? sourceRuns[sourceRuns.length - 1] ?? { text: '' };
      runs.push({
        ...run,
        text: run.text.includes('\u00a0') ? text : text.replace(/\u00a0/g, ' '),
      });
    };
    const visit = (node: Node): void => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = stripZwsp(node.textContent ?? '');
        if (text) append(text, Math.max(0, index - 1));
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        if (element.dataset.bullet === 'true') return;
        if (element.hasAttribute('data-pptx-run')) {
          const sourceIndex = Number(element.dataset.pptxRun);
          append(runText(element), sourceIndex);
          index = sourceIndex + 1;
        } else {
          for (const child of Array.from(node.childNodes)) visit(child);
        }
      }
    };
    for (const child of Array.from(paragraph.childNodes)) visit(child);
    return {
      ...source,
      runs: runs.length ? runs : [{ ...(sourceRuns[0] ?? {}), text: '' }],
    } as P;
  });
}
