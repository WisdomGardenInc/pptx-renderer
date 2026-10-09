import type { PresentationData } from '../model/Presentation';
import { materializeSlideNodes } from '../model/Presentation';
import type { SlideData } from '../model/Slide';
import { parseSlide } from '../model/Slide';
import { parseShapeNode } from '../model/nodes/ShapeNode';
import { parseXml } from '../parser/XmlParser';
import { renderSlide } from '../renderer/SlideRenderer';
import type { SlideHandle, SlideRendererOptions } from '../renderer/SlideRenderer';
import { editSlideElements } from '../writer/ElementWriter';
import type { EditableElement } from './EditableElement';
import { sourceElement } from './parts';
import { readSlideElements } from './readSlideElements';
import { createEditableTextRange, mapEditableTextPositionToParagraph } from './editableTextDom';

export interface EditableTextHandle {
  /** Transparent overlay in native 96-DPI pixels. The host applies its canvas scale. */
  readonly element: HTMLDivElement;
  /** Stable input root, retained by update() so listeners and IME ownership survive. */
  readonly textElement: HTMLDivElement;
  readonly ready: Promise<void>;
  /** Apply model text/geometry without mutating the presentation or slide XML. */
  update(element: EditableElement, context?: EditableTextContext): void;
  dispose(): void;
}

export interface EditableTextContext {
  readonly presentation: PresentationData;
  readonly slide: SlideData;
}

export type EditableTextOptions = Pick<
  SlideRendererOptions,
  'mediaUrlCache' | 'fontFaces' | 'embeddedFonts' | 'embeddedFontLimits' | 'pdfjs'
>;

/** Render one editable text/shape using the same inherited layout as the slide preview. */
export function renderEditableText(
  presentation: PresentationData,
  slide: SlideData,
  initial: EditableElement,
  options: EditableTextOptions = {},
): EditableTextHandle {
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
    // ElementWriter mutates its tree. Parse an independent slide for every edit
    // so temporary input cannot change other objects, shape order or saved XML.
    const xml =
      slide.root?.exists() && slide.root.element
        ? new XMLSerializer().serializeToString(slide.root.element)
        : slide.sourceXml;
    if (!xml) throw new Error('Missing PPTX slide XML.');
    const temporary = parseSlide(
      parseXml(xml),
      slide.index,
      new Map(slide.rels),
      slide.slidePath,
      presentation.diagramDrawings,
    );
    materializeSlideNodes(presentation, temporary);
    const original = temporary.nodes.find((node) => node.id === element.nodeId);
    const previous = readSlideElements(presentation, temporary).find(
      (item) => item.nodeId === element.nodeId,
    );
    if (original?.nodeType !== 'shape' || !previous) throw new Error('Missing PPTX text node.');
    editSlideElements(temporary, [element], [{ ...previous, id: element.id }], null, new Map());
    const source = sourceElement(temporary, element.nodeId);
    const shape =
      source?.localName === 'sp'
        ? source
        : source?.getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/presentationml/2006/main',
            'sp',
          )[0];
    if (!shape) throw new Error('Missing PPTX text shape.');
    const parsed = parseShapeNode(parseXml(new XMLSerializer().serializeToString(shape)));
    const node = {
      ...original,
      source: parsed.source,
      position: parsed.position,
      size: parsed.size,
      rotation: parsed.rotation,
      flipH: parsed.flipH,
      flipV: parsed.flipV,
      textBody: parsed.textBody
        ? {
            ...parsed.textBody,
            layoutBodyProperties: original.textBody?.layoutBodyProperties,
          }
        : undefined,
    };
    const background = parseXml(
      '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:bgPr><a:noFill/></p:bgPr></p:bg>',
    );
    const current = input;
    const selection = current?.ownerDocument.getSelection();
    const focused = !!current && current.ownerDocument.activeElement === current;
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
      { ...temporary, nodes: [node], background, showMasterSp: false },
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
  return {
    element: overlay,
    get textElement() {
      return input!;
    },
    get ready() {
      return view!.ready;
    },
    update,
    dispose() {
      if (disposed) return;
      disposed = true;
      view?.dispose();
      if (!options.mediaUrlCache) {
        for (const url of media.values()) if (url.startsWith('blob:')) URL.revokeObjectURL(url);
        media.clear();
      }
    },
  };
}
