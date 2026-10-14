import type { ShapeNodeData, TextBody } from '../model/nodes/ShapeNode';
import { SafeXmlNode } from '../parser/XmlParser';
import { getEffectiveBodyPrChild, parseTextPercentage } from '../renderer/TextBodyProperties';
import { A, createEl, insertOrdered, removeChildrenByLocal } from '../writer/xmlEdit';

export type TextAutoFit = 'none' | 'shrink' | 'resize';

export interface TextFitMetrics {
  fontScale?: number | null;
  lineSpaceReduction?: number | null;
}

export interface TextLayout {
  autoFit: TextAutoFit | null;
  wrap: boolean;
  fontScale: number | null;
  lineSpaceReduction: number | null;
}

export interface TextLayoutPatch extends TextFitMetrics {
  autoFit?: TextAutoFit | null;
  wrap?: boolean;
}

interface TextFitProperties extends TextFitMetrics {
  autoFit?: TextAutoFit | null;
}

const AUTO_FIT_ELEMENTS = {
  none: 'noAutofit',
  shrink: 'normAutofit',
  resize: 'spAutoFit',
} as const;

const FIT_METRIC_ATTRIBUTES = {
  fontScale: 'fontScale',
  lineSpaceReduction: 'lnSpcReduction',
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
  const fit = getEffectiveBodyPrChild(textBody, 'normAutofit');
  const wrap =
    textBody?.bodyProperties?.attr('wrap') ?? textBody?.layoutBodyProperties?.attr('wrap');
  return {
    autoFit: autoFit ?? null,
    wrap: wrap !== 'none',
    fontScale: parseTextPercentage(fit?.attr('fontScale')) ?? null,
    lineSpaceReduction: parseTextPercentage(fit?.attr('lnSpcReduction')) ?? null,
  };
}

export function writeTextAutoFit(
  properties: Element,
  mode: TextAutoFit | null,
  metrics: TextFitMetrics = {},
): void {
  const previous = new SafeXmlNode(properties).child('normAutofit');
  const attributes = Object.entries(FIT_METRIC_ATTRIBUTES).map(([key, attribute]) => {
    const value = metrics[key as keyof TextFitMetrics];
    return [
      attribute,
      value === undefined ? parseTextPercentage(previous.attr(attribute)) : value,
    ] as const;
  });
  removeChildrenByLocal(properties, Object.values(AUTO_FIT_ELEMENTS));
  if (mode === null) return;
  const fit = createEl(properties.ownerDocument, A, `a:${AUTO_FIT_ELEMENTS[mode]}`);
  if (mode === 'shrink') {
    for (const [attribute, value] of attributes) {
      if (value != null) fit.setAttribute(attribute, String(Math.round(value * 100000)));
    }
  }
  insertOrdered(properties, fit, BODY_PR_ORDER);
}

export function writeTextFitProperties(
  properties: Element,
  current: TextFitProperties,
  previous?: TextFitProperties,
): void {
  const changed = (['autoFit', 'fontScale', 'lineSpaceReduction'] as const).some(
    (key) => current[key] !== undefined && current[key] !== previous?.[key],
  );
  if (!changed) return;
  const mode = current.autoFit === undefined ? previous?.autoFit : current.autoFit;
  if (mode === undefined) return;
  writeTextAutoFit(properties, mode, current);
}

export function applyTextLayout(shape: ShapeNodeData, patch: TextLayoutPatch): void {
  const doc =
    shape.source.element?.ownerDocument ?? document.implementation.createDocument(A, null);
  const body = (shape.textBody ??= { paragraphs: [] });
  const properties = body.bodyProperties?.element?.cloneNode(true) as Element | undefined;
  const next = properties ?? createEl(doc, A, 'a:bodyPr');
  if (patch.wrap !== undefined) next.setAttribute('wrap', patch.wrap ? 'square' : 'none');
  if (patch.autoFit !== undefined) {
    writeTextAutoFit(next, patch.autoFit, {
      fontScale: patch.fontScale ?? null,
      lineSpaceReduction: patch.lineSpaceReduction ?? null,
    });
  } else if (patch.fontScale !== undefined || patch.lineSpaceReduction !== undefined) {
    const layout = readTextLayout(body);
    writeTextAutoFit(next, layout.autoFit, { ...layout, ...patch });
  }
  body.bodyProperties = new SafeXmlNode(next);
}
