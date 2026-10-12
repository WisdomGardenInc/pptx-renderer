import type { SlideData } from '../model/Slide';
import { resolveRelTarget } from '../parser/RelParser';

export function relationshipsPart(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`;
}

export function resolvePartPath(source: string, target: string): string {
  return resolveRelTarget(source.slice(0, source.lastIndexOf('/')), target);
}

export function sourceElement(slide: SlideData, id: string): Element | undefined {
  const tree = slide.root?.child('cSld').child('spTree');
  if (!tree?.exists() || !tree.element) return undefined;
  const namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const find = (parent: Element): Element | undefined => {
    for (const element of Array.from(parent.children)) {
      if (['nvGrpSpPr', 'grpSpPr'].includes(element.localName)) continue;
      if (element.getElementsByTagNameNS(namespace, 'cNvPr')[0]?.getAttribute('id') === id)
        return element;
      if (
        element.localName === 'grpSp' ||
        element.localName === 'AlternateContent' ||
        element.localName === 'Choice' ||
        element.localName === 'Fallback'
      ) {
        const descendant = find(element);
        if (descendant) return descendant;
      }
    }
    return undefined;
  };
  return find(tree.element);
}
