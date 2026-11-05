import type { SafeXmlNode } from '../parser/XmlParser';
import type { PlaceholderInfo } from './nodes/BaseNode';
import type { LayoutData, PlaceholderEntry } from './Layout';
import type { MasterData } from './Master';

/** Read placeholder identity across all supported non-visual property wrappers. */
export function getPlaceholderInfo(node: SafeXmlNode): PlaceholderInfo {
  for (const wrapper of ['nvSpPr', 'nvPicPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr']) {
    const ph = node.child(wrapper).child('nvPr').child('ph');
    if (ph.exists()) return { type: ph.attr('type'), idx: ph.numAttr('idx') };
  }
  return {};
}

/** Slide -> layout matches idx (default zero); layout -> master matches category, never idx. */
export function resolvePlaceholderParents(
  placeholder: PlaceholderInfo | undefined,
  layout: LayoutData | undefined,
  master: MasterData | undefined,
): { layout?: PlaceholderEntry; master?: PlaceholderEntry; category: 'title' | 'body' | 'other' } {
  if (!placeholder) return { category: 'other' };
  const layoutMatch = layout?.placeholders.find(
    (entry) => (getPlaceholderInfo(entry.node).idx ?? 0) === (placeholder.idx ?? 0),
  );
  const type = masterPlaceholderType(
    layoutMatch ? getPlaceholderInfo(layoutMatch.node).type : placeholder.type,
  );
  const masterMatch = (
    master?.placeholderEntries ?? master?.placeholders.map((node) => ({ node }))
  )?.find((entry) => (getPlaceholderInfo(entry.node).type ?? 'obj') === type);
  const category =
    type === 'title' ? 'title' : ['body', 'dt', 'ftr', 'sldNum'].includes(type) ? 'body' : 'other';
  return { layout: layoutMatch, master: masterMatch, category };
}

/**
 * Layout placeholder categories that inherit from the master's body/title placeholder.
 * Master inheritance uses this type mapping, not the layout/slide idx namespace.
 * Reference: https://python-pptx.readthedocs.io/en/develop/dev/analysis/placeholders/layout-placeholders.html
 */
export function masterPlaceholderType(type: string | undefined): string {
  const effectiveType = type ?? 'obj';
  if (effectiveType === 'ctrTitle') return 'title';
  if (
    ['obj', 'subTitle', 'pic', 'chart', 'clipArt', 'dgm', 'media', 'tbl'].includes(effectiveType)
  ) {
    return 'body';
  }
  return effectiveType;
}
