import type { SafeXmlNode } from '../parser/XmlParser';
import type { EditableElement, EditableLineEnd } from './EditableElement';

const END_TYPES = new Set(['none', 'triangle', 'arrow', 'stealth', 'diamond', 'oval']);
const END_SIZES = new Set(['sm', 'med', 'lg']);

function paintAlpha(properties: SafeXmlNode): number | undefined {
  if (properties.child('noFill').exists()) return 0;
  const fill = properties.child('solidFill');
  if (!fill.exists()) return undefined;
  const color = fill.allChildren()[0];
  return (color?.child('alpha').numAttr('val') ?? 100000) / 100000;
}

function lineEnd(properties: SafeXmlNode): EditableLineEnd | undefined {
  const type = properties.attr('type');
  if (!type || !END_TYPES.has(type)) return undefined;
  const w = properties.attr('w');
  const len = properties.attr('len');
  return {
    type: type as EditableLineEnd['type'],
    ...(w && END_SIZES.has(w) ? { w: w as EditableLineEnd['w'] } : {}),
    ...(len && END_SIZES.has(len) ? { len: len as EditableLineEnd['len'] } : {}),
  };
}

export function readShapeEditingStyle(
  properties: SafeXmlNode,
  width: number,
  height: number,
): Partial<EditableElement> {
  const stroke = properties.child('ln');
  const geometry = properties.child('prstGeom');
  let cornerRadius: number | undefined;
  if (geometry.attr('prst') === 'roundRect') {
    const adjustment = geometry
      .child('avLst')
      .children('gd')
      .find((node) => node.attr('name') === 'adj');
    const value = Number(adjustment?.attr('fmla')?.replace(/^val\s+/, '') ?? 16667);
    const radius = Number.isFinite(value) ? value : 16667;
    cornerRadius =
      (Math.max(0, Math.min(width, height)) * Math.max(0, Math.min(50000, radius))) / 100000;
  }
  const cap = stroke.attr('cap');
  const lineCap = cap === 'rnd' || cap === 'sq' || cap === 'flat' ? cap : undefined;
  const headEnd = lineEnd(stroke.child('headEnd'));
  const tailEnd = lineEnd(stroke.child('tailEnd'));
  const fillAlpha = paintAlpha(properties);
  const strokeAlpha = paintAlpha(stroke);
  const dash = stroke.child('prstDash').attr('val');
  const strokeDash = (
    { solid: 'solid', dash: 'dashed', dot: 'dotted', dashDot: 'dashdot' } as const
  )[dash as 'solid'];
  return {
    ...(cornerRadius !== undefined ? { cornerRadius } : {}),
    ...(lineCap !== undefined ? { lineCap } : {}),
    ...(headEnd !== undefined ? { headEnd } : {}),
    ...(tailEnd !== undefined ? { tailEnd } : {}),
    ...(fillAlpha !== undefined ? { fillAlpha } : {}),
    ...(strokeAlpha !== undefined ? { strokeAlpha } : {}),
    ...(strokeDash !== undefined ? { strokeDash } : {}),
  };
}
