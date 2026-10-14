import type { ShapeNodeData, TextBody } from '../model/nodes/ShapeNode';
import { SafeXmlNode } from '../parser/XmlParser';
import { getEffectiveBodyPrChild } from '../renderer/TextBodyProperties';
import { A, createEl, insertOrdered, removeChildrenByLocal } from '../writer/xmlEdit';

export type TextAutoFit = 'none' | 'shrink' | 'resize';

export interface TextLayout {
  autoFit: TextAutoFit | null;
  wrap: boolean;
}

export interface TextLayoutPatch {
  autoFit?: TextAutoFit | null;
  wrap?: boolean;
}

const AUTO_FIT_ELEMENTS = {
  none: 'noAutofit',
  shrink: 'normAutofit',
  resize: 'spAutoFit',
} as const;

const BODY_PR_ORDER = [
  'prstTxWarp',
  'noAutofit',
  'normAutofit',
  'spAutoFit',
  'scene3d',
  'sp3d',
  'flatTx',
  'extLst',
];

export function readTextLayout(textBody: TextBody | undefined): TextLayout {
  const autoFit = (Object.keys(AUTO_FIT_ELEMENTS) as TextAutoFit[]).find((mode) =>
    getEffectiveBodyPrChild(textBody, AUTO_FIT_ELEMENTS[mode])?.exists(),
  );
  const wrap =
    textBody?.bodyProperties?.attr('wrap') ?? textBody?.layoutBodyProperties?.attr('wrap');
  return { autoFit: autoFit ?? null, wrap: wrap !== 'none' };
}

export function writeTextAutoFit(properties: Element, mode: TextAutoFit | null): void {
  removeChildrenByLocal(properties, Object.values(AUTO_FIT_ELEMENTS));
  if (mode === null) return;
  const fit = createEl(properties.ownerDocument, A, `a:${AUTO_FIT_ELEMENTS[mode]}`);
  insertOrdered(properties, fit, BODY_PR_ORDER);
}

export function applyTextLayout(shape: ShapeNodeData, patch: TextLayoutPatch): void {
  const doc =
    shape.source.element?.ownerDocument ?? document.implementation.createDocument(A, null);
  const body = (shape.textBody ??= { paragraphs: [] });
  const properties = body.bodyProperties?.element?.cloneNode(true) as Element | undefined;
  const next = properties ?? createEl(doc, A, 'a:bodyPr');
  if (patch.wrap !== undefined) next.setAttribute('wrap', patch.wrap ? 'square' : 'none');
  if (patch.autoFit !== undefined) writeTextAutoFit(next, patch.autoFit);
  body.bodyProperties = new SafeXmlNode(next);
}
