import type { PresentationData } from '../model/Presentation';
import { materializeSlideNodes, resolveNodePlaceholderInheritance } from '../model/Presentation';
import type { SlideData } from '../model/Slide';
import { parseShapeNode } from '../model/nodes/ShapeNode';
import { parseXml, SafeXmlNode } from '../parser/XmlParser';
import { createRenderContext } from '../renderer/RenderContext';
import { renderSlide } from '../renderer/SlideRenderer';
import type { SlideHandle, SlideRendererOptions } from '../renderer/SlideRenderer';
import { patchTextElement } from '../writer/ElementWriter';
import type { EditableElement } from './EditableElement';
import { findSlideElementNode, groupLocalElement, mapSlideElementNode } from './slideElementNodes';
import type { SlideElementNode } from './slideElementNodes';
import { projectSlideElement } from './readSlideElements';
import { measureTextContentSize, measureEditableTextSize } from './measureEditableTextSize';
import { readTextLayout } from './textLayout';
import {
  captureEditableWhitespace,
  createEditableTextRange,
  extractEditableText,
  normalizeEditableWhitespace,
  restoreEditableWhitespace,
  mapEditableTextPositionToParagraph,
} from './editableTextDom';

export interface EditableTextHandle {
  /** Transparent overlay in native 96-DPI pixels. The host applies its canvas scale. */
  readonly element: HTMLDivElement;
  /** Stable input root, retained by update() so listeners and IME ownership survive. */
  readonly textElement: HTMLDivElement;
  readonly ready: Promise<void>;
  extract(element?: EditableElement, options?: EditableTextExtractOptions): EditableElement;
  /** Apply model text/geometry without mutating the presentation or slide XML. */
  update(element: EditableElement, context?: EditableTextContext): void;
  dispose(): void;
}

export interface EditableTextExtractOptions {
  readonly fitToContent?: boolean;
}

export interface EditableTextContext {
  readonly presentation: PresentationData;
  readonly slide: SlideData;
}

export type EditableTextOptions = Pick<
  SlideRendererOptions,
  | 'mediaUrlCache'
  | 'fontFaces'
  | 'embeddedFonts'
  | 'embeddedFontLimits'
  | 'pdfjs'
  | 'editableTextAutoFit'
>;

function textWasReduced(next: EditableElement, previous: EditableElement): boolean {
  const text = (element: EditableElement) =>
    (element.paragraphs ?? [])
      .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
      .join('\n');
  const before = text(previous);
  const after = text(next);
  return after.length < before.length || after.split('\n').length < before.split('\n').length;
}

function textLayoutSignature(element: EditableElement): string {
  return JSON.stringify([
    element.paragraphs,
    element.wrap,
    element.autoFit,
    element.width,
    element.bodyPadding,
  ]);
}

export function renderEditableText(
  presentation: PresentationData,
  slide: SlideData,
  initial: EditableElement,
  options: EditableTextOptions = {},
): EditableTextHandle {
  return createEditableTextHandle(presentation, slide, initial, options);
}

/** Internal batch entry point: fitting already has each node and its draw order. */
export function renderEditableTextFromNode(
  presentation: PresentationData,
  slide: SlideData,
  initial: EditableElement,
  match: { entry: SlideElementNode; index: number },
  options: EditableTextOptions,
): EditableTextHandle {
  return createEditableTextHandle(presentation, slide, initial, options, match);
}

function createEditableTextHandle(
  presentation: PresentationData,
  slide: SlideData,
  initial: EditableElement,
  options: EditableTextOptions,
  initialMatch?: { entry: SlideElementNode; index: number },
): EditableTextHandle {
  const background = parseXml(
    '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:bgPr><a:noFill/></p:bgPr></p:bg>',
  );
  const overlay = document.createElement('div');
  Object.assign(overlay.style, {
    position: 'absolute',
    top: '0',
    left: '0',
    width: `${presentation.width}px`,
    height: `${presentation.height}px`,
    pointerEvents: 'none',
    overflow: 'visible',
  });
  const media = options.mediaUrlCache ?? new Map<string, string>();
  let view: SlideHandle | undefined;
  let input: HTMLDivElement | undefined;
  let renderedElement = initial;
  let renderedSize = { width: initial.width, height: initial.height };
  let renderedTextSize = { ...renderedSize };
  let fitContent = false;
  let resizeToContent = false;
  let fixedFrame = false;
  let resizeContentChanged = false;
  let wrapsText = true;
  let disposed = false;

  const update = (element: EditableElement, context?: EditableTextContext): void => {
    if (context) {
      presentation = context.presentation;
      slide = context.slide;
    }
    if (disposed) throw new Error('The editable text handle has been disposed.');
    if (!element.nodeId || (element.type !== 'text' && element.type !== 'shape')) {
      throw new Error('An existing PPTX text or shape node is required.');
    }
    // Clone only the target shape. Its native source stays detached from temporary
    // edits; group coordinate mapping and the slide's theme context remain intact.
    materializeSlideNodes(presentation, slide);
    const match =
      initialMatch?.entry.node.id === element.nodeId
        ? initialMatch
        : findSlideElementNode(slide, element.nodeId, presentation);
    initialMatch = undefined;
    if (!match || match.entry.node.nodeType !== 'shape' || !match.entry.node.source.exists())
      throw new Error('Missing PPTX text shape.');
    const ctx = createRenderContext(presentation, slide);
    const previous = projectSlideElement(
      presentation,
      slide,
      match.entry,
      match.index,
      { includeSourceXml: false },
      ctx,
    );
    const source = match.entry.node.source.element!.cloneNode(true) as Element;
    const { node: original, parent } = match.entry;
    patchTextElement(
      source,
      parent ? groupLocalElement(element, original, parent) : element,
      parent ? groupLocalElement(previous, original, parent) : previous,
    );
    const localNode = parseShapeNode(new SafeXmlNode(source));
    resolveNodePlaceholderInheritance(localNode, ctx.layout, ctx.master);
    const node = mapSlideElementNode(localNode, parent).node;
    if (node.nodeType !== 'shape') throw new Error('Missing PPTX text shape.');
    const current = input;
    const selection = current?.ownerDocument.getSelection();
    const focused = !!current && current.ownerDocument.activeElement === current;
    const whitespace =
      focused && element.nodeId === renderedElement.nodeId
        ? captureEditableWhitespace(current, renderedElement.paragraphs ?? [])
        : undefined;
    const range = selection?.rangeCount ? selection.getRangeAt(0) : undefined;
    const start =
      range && current?.contains(range.startContainer)
        ? mapEditableTextPositionToParagraph(range.startContainer, range.startOffset, current!)
        : null;
    const end =
      range && current?.contains(range.endContainer)
        ? mapEditableTextPositionToParagraph(range.endContainer, range.endOffset, current!)
        : null;
    let native: HTMLElement | undefined;
    const next = renderSlide(
      presentation,
      {
        ...slide,
        nodes: [node],
        background,
        showMasterSp: false,
        nodesMaterialized: true,
        placeholderInheritanceResolved: true,
      },
      {
        ...options,
        mediaUrlCache: media,
        editableText: true,
        editableTextRoot: input,
        onNodeRendered: (candidate, rendered) => {
          if (candidate === node) native = rendered;
        },
      },
    );
    const text = native?.querySelector<HTMLDivElement>('[data-pptx-text-root]');
    if (!native || !text) {
      next.dispose();
      throw new Error('This PPTX shape has no editable text container.');
    }
    const root = text;
    if (whitespace) restoreEditableWhitespace(root, whitespace, element.paragraphs ?? []);
    const layout = readTextLayout(node.textBody);
    resizeToContent = options.editableTextAutoFit === true && layout.autoFit === 'resize';
    fixedFrame =
      options.editableTextAutoFit === true &&
      (layout.autoFit === 'none' || layout.autoFit === 'shrink');
    wrapsText = layout.wrap;
    if (element.nodeId !== renderedElement.nodeId) {
      fitContent = false;
      resizeContentChanged = false;
    } else {
      if (textWasReduced(element, renderedElement)) fitContent = true;
      if (textLayoutSignature(element) !== textLayoutSignature(renderedElement))
        resizeContentChanged = true;
    }
    if (fitContent && root.style.height === 'auto') {
      root.style.minHeight = '0px';
      if (root.style.width === 'max-content') root.style.minWidth = '0px';
    }
    root.setAttribute('contenteditable', 'true');
    root.tabIndex = 0;
    root.style.pointerEvents = 'auto';
    root.style.outline = 'none';
    root.style.cursor = 'text';
    input = root;
    next.element.replaceChildren(native);
    next.element.style.background = 'none';
    next.element.style.overflow = 'visible';
    overlay.replaceChildren(next.element);
    const old = view;
    view = next;
    renderedElement = element;
    renderedSize = { width: node.size.w, height: node.size.h };
    renderedTextSize = {
      width: node.textBoxBounds?.w ?? node.size.w,
      height: node.textBoxBounds?.h ?? node.size.h,
    };
    old?.dispose();
    if (focused) {
      root.focus({ preventScroll: true });
      if (start && end) {
        const restored = createEditableTextRange(
          root,
          start.paraIdx,
          start.flatChar,
          end.paraIdx,
          end.flatChar,
        );
        if (restored) {
          selection?.removeAllRanges();
          selection?.addRange(restored);
        }
      }
    }
  };

  update(initial);
  const root = input!;
  let composing = false;
  const normalizeWhitespace = () => {
    if (!disposed && !composing)
      normalizeEditableWhitespace(root, renderedElement.paragraphs ?? []);
  };
  const startComposition = () => {
    composing = true;
  };
  const endComposition = () => {
    composing = false;
    normalizeWhitespace();
  };
  const Observer = root.ownerDocument.defaultView?.MutationObserver;
  const whitespaceObserver =
    options.editableTextAutoFit && Observer ? new Observer(normalizeWhitespace) : undefined;
  if (options.editableTextAutoFit) {
    root.addEventListener('input', normalizeWhitespace);
    root.addEventListener('compositionstart', startComposition);
    root.addEventListener('compositionend', endComposition);
    whitespaceObserver?.observe(root, { childList: true, subtree: true, characterData: true });
  }
  return {
    element: overlay,
    get textElement() {
      return input!;
    },
    get ready() {
      return view!.ready;
    },
    extract(element = renderedElement, extraction = {}) {
      if (disposed) throw new Error('The editable text handle has been disposed.');
      const root = input!;
      const paragraphs = extractEditableText(root, element.paragraphs ?? []);
      if (textWasReduced({ ...element, paragraphs }, renderedElement)) fitContent = true;
      if (JSON.stringify(paragraphs) !== JSON.stringify(renderedElement.paragraphs))
        resizeContentChanged = true;
      const shouldFitContent = resizeToContent && (resizeContentChanged || extraction.fitToContent);
      let size = { width: element.width, height: element.height };
      if (!fixedFrame && root.isConnected) {
        size = shouldFitContent
          ? measureTextContentSize(root, renderedSize, renderedTextSize, wrapsText)
          : measureEditableTextSize(root, element, renderedSize, renderedTextSize, fitContent);
      }
      return { ...element, ...size, paragraphs };
    },
    update,
    dispose() {
      if (disposed) return;
      disposed = true;
      whitespaceObserver?.disconnect();
      root.removeEventListener('input', normalizeWhitespace);
      root.removeEventListener('compositionstart', startComposition);
      root.removeEventListener('compositionend', endComposition);
      view?.dispose();
      if (!options.mediaUrlCache) {
        for (const url of media.values()) if (url.startsWith('blob:')) URL.revokeObjectURL(url);
        media.clear();
      }
    },
  };
}
