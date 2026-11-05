import type { SlideData } from '../model/Slide';
import { resolveRelTarget } from '../parser/RelParser';

export function relationshipsPart(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`;
}

export function resolvePartPath(source: string, target: string): string {
  return resolveRelTarget(source.slice(0, source.lastIndexOf('/')), target);
}

/** Build one source lookup per operation; callers rebuild it after mutating the tree. */
export function indexSourceElements(slide: SlideData): Map<string, Element> {
  const sources = new Map<string, Element>();
  const tree = slide.root?.child('cSld').child('spTree');
  if (!tree?.exists() || !tree.element) return sources;
  const namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const firstIds = new WeakMap<Element, string>();
  const collect = (element: Element): string | undefined => {
    let first =
      element.namespaceURI === namespace && element.localName === 'cNvPr'
        ? (element.getAttribute('id') ?? undefined)
        : undefined;
    for (const child of Array.from(element.children)) {
      const id = collect(child);
      first ??= id;
    }
    if (first !== undefined) firstIds.set(element, first);
    return first;
  };
  collect(tree.element);
  const visit = (parent: Element): void => {
    for (const element of Array.from(parent.children)) {
      if (['nvGrpSpPr', 'grpSpPr'].includes(element.localName)) continue;
      const id = firstIds.get(element);
      if (id !== undefined && !sources.has(id)) sources.set(id, element);
      if (['grpSp', 'AlternateContent', 'Choice', 'Fallback'].includes(element.localName))
        visit(element);
    }
  };
  visit(tree.element);
  return sources;
}

export function sourceElement(
  slide: SlideData,
  id: string,
  sources = indexSourceElements(slide),
): Element | undefined {
  return sources.get(id);
}
