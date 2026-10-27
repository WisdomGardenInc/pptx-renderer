import { parseXml } from '../parser/XmlParser';
import type { EditableLineEnd } from '../editor/EditableElement';
import {
  A,
  SP_PR_ORDER,
  RPR_CHILD_ORDER,
  FILL_LOCAL_NAMES,
  childByLocal,
  ensureChild,
  ensureShapeProperties,
  removeChildrenByLocal,
} from './xmlEdit';
import { LINE_STYLE_ORDER, patchLineEnd } from './shapeEditingStyle';

export interface ShapeFillPaint {
  readonly color: string | null;
  readonly alpha?: number;
}
export interface ShapeStrokePaint {
  readonly color?: string | null;
  readonly alpha?: number;
  readonly width?: number;
  readonly cap?: 'flat' | 'rnd' | 'sq';
  readonly join?: 'round' | 'bevel' | 'miter';
  readonly dash?: 'solid' | 'dashed' | 'dotted' | 'dashdot';
  readonly headEnd?: EditableLineEnd | null;
  readonly tailEnd?: EditableLineEnd | null;
}
export interface ShapePaint {
  readonly fill?: ShapeFillPaint;
  readonly stroke?: ShapeStrokePaint;
}

export function setFillColor(
  parent: Element,
  color: string | null | undefined,
  alpha?: number | null,
): void {
  removeChildrenByLocal(parent, FILL_LOCAL_NAMES);
  let order: readonly string[] = SP_PR_ORDER;
  if (parent.localName === 'ln') order = LINE_STYLE_ORDER;
  if (parent.localName === 'rPr') order = RPR_CHILD_ORDER;
  const fill = ensureChild(parent, A, color ? 'a:solidFill' : 'a:noFill', order);
  if (!color) return;
  const rgb = ensureChild(fill, A, 'a:srgbClr', []);
  rgb.setAttribute('val', color.replace(/^#/, ''));
  if (alpha != null && alpha < 1)
    ensureChild(rgb, A, 'a:alpha', []).setAttribute('val', String(Math.round(alpha * 100000)));
}

const COLOR_LOCAL_NAMES = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr', 'scrgbClr'];

/** Set absolute opacity without replacing the authored fill or color transforms. */
export function setFillAlpha(
  source: Element,
  alpha: number | null | undefined,
  sourceFillXml?: string,
): void {
  const value = alpha ?? 1;
  if (!Number.isFinite(value) || value < 0 || value > 1)
    throw new RangeError('Paint alpha must be between zero and one.');
  const properties = childByLocal(source, 'spPr');
  if (!properties) return;
  let explicit = Array.from(properties.children).find((node) =>
    FILL_LOCAL_NAMES.includes(node.localName),
  );
  if (!explicit && sourceFillXml) {
    const inherited = parseXml(sourceFillXml);
    if (inherited.exists() && FILL_LOCAL_NAMES.includes(inherited.localName)) {
      // fillRef alpha alone is ignored by some Office consumers. Materialize the
      // effective theme paint locally while retaining its scheme colors and geometry.
      explicit = properties.ownerDocument.importNode(inherited.element!, true) as Element;
      ensureChild(properties, A, `a:${inherited.localName}`, SP_PR_ORDER).replaceWith(explicit);
    }
  }
  const reference = childByLocal(childByLocal(source, 'style') ?? source, 'fillRef');
  const fill = explicit ?? reference;
  if (!fill || fill.localName === 'noFill') return;
  if (!explicit && Number(reference?.getAttribute('idx') ?? 0) <= 0) return;
  if (fill.localName === 'blipFill') {
    const blip = childByLocal(fill, 'blip');
    if (blip) {
      removeChildrenByLocal(blip, ['alphaModFix', 'alphaMod', 'alphaOff']);
      if (value < 1) {
        const modifier = blip.ownerDocument.createElementNS(A, 'a:alphaModFix');
        modifier.setAttribute('amt', String(Math.round(value * 100000)));
        blip.insertBefore(modifier, childByLocal(blip, 'extLst') ?? null);
      }
    }
    return;
  }
  for (const color of Array.from(fill.getElementsByTagNameNS(A, '*'))) {
    if (!COLOR_LOCAL_NAMES.includes(color.localName)) continue;
    removeChildrenByLocal(color, ['alpha', 'alphaMod', 'alphaOff']);
    if (value < 1)
      ensureChild(color, A, 'a:alpha', []).setAttribute('val', String(Math.round(value * 100000)));
  }
}

function validatePaint(paint: ShapePaint): void {
  for (const part of [paint.fill, paint.stroke]) {
    if (
      part?.alpha !== undefined &&
      (!Number.isFinite(part.alpha) || part.alpha < 0 || part.alpha > 1)
    )
      throw new RangeError('Paint alpha must be between zero and one.');
  }
  const width = paint.stroke?.width;
  if (width !== undefined && (!Number.isFinite(width) || width < 0))
    throw new RangeError('Stroke width must be finite and nonnegative.');
}

export function writeShapePaint(source: Element, paint: ShapePaint): void {
  validatePaint(paint);
  const properties = ensureShapeProperties(source);
  if (paint.fill) setFillColor(properties, paint.fill.color, paint.fill.alpha);
  const style = paint.stroke;
  if (!style) return;
  const stroke = ensureChild(properties, A, 'a:ln', SP_PR_ORDER);
  if (style.color !== undefined) setFillColor(stroke, style.color, style.alpha);
  if (style.width !== undefined) stroke.setAttribute('w', String(Math.round(style.width * 9525)));
  if (style.cap !== undefined) stroke.setAttribute('cap', style.cap);
  if (style.join !== undefined) {
    removeChildrenByLocal(stroke, ['round', 'bevel', 'miter']);
    ensureChild(stroke, A, `a:${style.join}`, LINE_STYLE_ORDER);
  }
  if (style.dash !== undefined) {
    childByLocal(stroke, 'custDash')?.remove();
    ensureChild(stroke, A, 'a:prstDash', LINE_STYLE_ORDER).setAttribute(
      'val',
      { solid: 'solid', dashed: 'dash', dotted: 'dot', dashdot: 'dashDot' }[style.dash],
    );
  }
  for (const name of ['headEnd', 'tailEnd'] as const) {
    const end = style[name];
    if (end !== undefined) patchLineEnd(stroke, name, end);
  }
}
