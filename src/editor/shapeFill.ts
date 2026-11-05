import type { EditableElement } from './EditableElement';
import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import type { SafeXmlNode } from '../parser/XmlParser';
import type { RenderContext } from '../renderer/RenderContext';
import { resolveColorWithPlaceholder } from '../renderer/StyleResolver';
import { inheritedGroupFill } from '../model/groupFill';
import { SafeXmlNode as XmlNode } from '../parser/XmlParser';

type FillStyle = Pick<EditableElement, 'fillKind' | 'fillColor' | 'fillAlpha' | 'sourceFillXml'>;

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function boundThemeFill(fill: SafeXmlNode, reference: SafeXmlNode): string | undefined {
  if (!fill.exists() || !reference.exists()) return undefined;
  const color = reference.allChildren()[0];
  const copy =
    fill.localName === 'fillRef'
      ? fill.element!.ownerDocument.createElementNS(A, 'a:solidFill')
      : (fill.element!.cloneNode(true) as Element);
  if (fill.localName === 'fillRef' && color?.exists())
    copy.appendChild(color.element!.cloneNode(true));
  for (const placeholder of Array.from(copy.getElementsByTagNameNS(A, 'schemeClr'))) {
    if (placeholder.getAttribute('val')?.toLowerCase() !== 'phclr' || !color?.exists()) continue;
    const bound = color.element!.cloneNode(true) as Element;
    // Reference modifiers run first, followed by modifiers on the theme placeholder.
    for (const modifier of Array.from(placeholder.children))
      bound.appendChild(modifier.cloneNode(true));
    placeholder.replaceWith(bound);
  }
  return new XMLSerializer().serializeToString(copy);
}

/** Project fill capabilities without flattening non-solid paint to a single color. */
export function readShapeFill(node: ShapeNodeData, ctx: RenderContext): FillStyle {
  let fill = node.fill;
  let placeholder: SafeXmlNode | undefined;
  let themed = false;
  let groupFillXml: string | undefined;
  if (fill?.localName === 'grpFill' && node.source.element) {
    const inherited = inheritedGroupFill(node.source.element);
    if (inherited) {
      fill = new XmlNode(inherited);
      groupFillXml = new XMLSerializer().serializeToString(inherited);
      themed = true;
    }
  }
  if (!fill?.exists()) {
    const reference = node.source.child('style').child('fillRef');
    const index = reference.numAttr('idx') ?? 0;
    if (!reference.exists() || index <= 0) return { fillKind: 'none', fillColor: null };
    themed = true;
    placeholder = reference;
    fill = ctx.theme.fillStyles[index - 1] ?? reference;
  }
  const kind = (
    {
      solidFill: 'solid',
      gradFill: 'gradient',
      pattFill: 'pattern',
      blipFill: 'image',
      grpFill: 'theme',
      noFill: 'none',
      fillRef: 'solid',
    } as Record<string, NonNullable<EditableElement['fillKind']>>
  )[fill.localName];
  const style: FillStyle = {
    fillKind: themed ? 'theme' : (kind ?? 'none'),
    fillColor: null,
    ...(placeholder ? { sourceFillXml: boundThemeFill(fill, placeholder) } : {}),
    ...(groupFillXml ? { sourceFillXml: groupFillXml } : {}),
  };
  if (kind === 'none') return { ...style, fillAlpha: 0 };
  if (kind === 'solid') {
    const resolved = resolveColorWithPlaceholder(fill, ctx, placeholder);
    return {
      ...style,
      fillColor: `#${resolved.color.replace(/^#/, '')}`,
      fillAlpha: resolved.alpha,
    };
  }
  const colors =
    kind === 'gradient'
      ? fill.child('gsLst').children('gs')
      : kind === 'pattern'
        ? [fill.child('fgClr'), fill.child('bgClr')].filter((color) => color.exists())
        : [];
  const alphas = colors.map((color) => resolveColorWithPlaceholder(color, ctx, placeholder).alpha);
  // A scalar cannot represent different per-stop alphas; leave it unset in that case.
  if (alphas.length && alphas.every((alpha) => alpha === alphas[0])) style.fillAlpha = alphas[0];
  if (kind === 'image') {
    const blip = fill.child('blip');
    style.fillAlpha = Math.max(
      0,
      Math.min(
        1,
        ((blip.child('alphaModFix').numAttr('amt') ?? 100000) / 100000) *
          ((blip.child('alphaMod').numAttr('val') ?? 100000) / 100000) +
          (blip.child('alphaOff').numAttr('val') ?? 0) / 100000,
      ),
    );
  }
  return style;
}
