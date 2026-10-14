import { materializeSlideNodes, type PresentationData } from '../model/Presentation';
import { parseSlide, type SlideData } from '../model/Slide';
import { parseRels } from '../parser/RelParser';
import { parseXml } from '../parser/XmlParser';
import { editSlideElements } from '../writer/ElementWriter';
import type { EditableElement } from './EditableElement';
import { readSlideElements } from './readSlideElements';
import { renderEditableText, type EditableTextOptions } from './renderEditableText';

export interface FitSlideTextOptions extends EditableTextOptions {
  readonly relationshipsXml?: string | null;
  readonly imageParts?: Map<string, string>;
}

function layoutSignature(element: EditableElement | undefined): string {
  if (!element) return '';
  return JSON.stringify([
    element.autoFit,
    element.fontScale,
    element.lineSpaceReduction,
    element.wrap,
    element.paragraphs,
    element.width,
    element.height,
    element.bodyPadding,
    element.vert,
  ]);
}

function relationshipsXml(slide: SlideData): string {
  const namespace = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const root = document.createElementNS(namespace, 'Relationships');
  for (const [id, entry] of slide.rels) {
    const relationship = document.createElementNS(namespace, 'Relationship');
    relationship.setAttribute('Id', id);
    relationship.setAttribute('Type', entry.type);
    relationship.setAttribute('Target', entry.target);
    if (entry.targetMode) relationship.setAttribute('TargetMode', entry.targetMode);
    root.appendChild(relationship);
  }
  return new XMLSerializer().serializeToString(root);
}

export async function fitSlideText<E extends EditableElement>(
  presentation: PresentationData,
  slide: SlideData,
  elements: readonly E[],
  previous: readonly EditableElement[],
  options: FitSlideTextOptions = {},
): Promise<E[]> {
  const previousById = new Map(previous.map((element) => [element.id, element]));
  const candidates = elements.filter((element) => {
    if (element.type !== 'text' && element.type !== 'shape') return false;
    const old = previousById.get(element.id);
    const mode = element.autoFit === undefined ? old?.autoFit : element.autoFit;
    return (
      mode !== 'none' && mode !== 'shrink' && layoutSignature(element) !== layoutSignature(old)
    );
  });
  if (candidates.length === 0) return [...elements];
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
  const part = editSlideElements(
    temporary,
    elements,
    previous,
    options.relationshipsXml ?? relationshipsXml(slide),
    options.imageParts ?? new Map(),
  );
  const prepared = parseSlide(
    parseXml(part.xml),
    slide.index,
    parseRels(part.relsXml ?? ''),
    slide.slidePath,
    presentation.diagramDrawings,
  );
  materializeSlideNodes(presentation, prepared);
  const projected = new Map(
    readSlideElements(presentation, prepared).map((element) => [element.nodeId, element]),
  );
  const nodeIds = new Map(Array.from(part.ids, ([nodeId, id]) => [id, nodeId]));
  const sizes = new Map<string, { width: number; height: number }>();
  for (const element of candidates) {
    const nodeId = nodeIds.get(element.id);
    if (!nodeId || projected.get(nodeId)?.autoFit !== 'resize') continue;
    const handle = renderEditableText(
      presentation,
      prepared,
      { ...element, nodeId },
      { ...options, editableTextAutoFit: true },
    );
    Object.assign(handle.element.style, {
      position: 'fixed',
      left: '-100000px',
      top: '0',
      visibility: 'hidden',
    });
    document.body.appendChild(handle.element);
    try {
      await handle.ready;
      const fitted = handle.extract(undefined, { fitToContent: true });
      sizes.set(element.id, { width: fitted.width, height: fitted.height });
    } finally {
      handle.dispose();
      handle.element.remove();
    }
  }
  return elements.map((element) => {
    const size = sizes.get(element.id);
    return size ? { ...element, ...size } : element;
  });
}
