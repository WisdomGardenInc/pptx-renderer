import { SafeXmlNode } from '../parser/XmlParser';
import type { TextBody, TextParagraph } from '../model/nodes/ShapeNode';
import type { PlaceholderInfo } from '../model/nodes/BaseNode';
import { resolvePlaceholderParents } from '../model/placeholderMatching';
import type { RenderContext } from './RenderContext';

/** Find the level-specific paragraph style, with defPPr as the fallback. */
export function findStyleAtLevel(style: SafeXmlNode | undefined, level: number): SafeXmlNode {
  if (!style?.exists()) return new SafeXmlNode(null);
  const properties = style.child(`lvl${level + 1}pPr`);
  return properties.exists() ? properties : style.child('defPPr');
}

/** Shared cascade in increasing precedence, used by rendering and editor projection. */
export function paragraphStyleSources(
  textBody: TextBody,
  placeholder: PlaceholderInfo | undefined,
  paragraph: TextParagraph,
  ctx: RenderContext,
): SafeXmlNode[] {
  const parents = resolvePlaceholderParents(placeholder, ctx.layout, ctx.master);
  const masterStyle =
    ctx.master.textStyles[
      parents.category === 'title'
        ? 'titleStyle'
        : parents.category === 'body'
          ? 'bodyStyle'
          : 'otherStyle'
    ];
  return [
    findStyleAtLevel(ctx.presentation.defaultTextStyle, paragraph.level),
    findStyleAtLevel(ctx.master.defaultTextStyle, paragraph.level),
    findStyleAtLevel(masterStyle, paragraph.level),
    findStyleAtLevel(parents.master?.node.child('txBody').child('lstStyle'), paragraph.level),
    findStyleAtLevel(parents.layout?.node.child('txBody').child('lstStyle'), paragraph.level),
    findStyleAtLevel(textBody.listStyle, paragraph.level),
    paragraph.properties,
  ].map((source) => source ?? new SafeXmlNode(null));
}
