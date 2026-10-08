import type { SlideData } from '../model/Slide';
import { resolveRelTarget } from '../parser/RelParser';

export function relationshipsPart(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`;
}

export function resolvePartPath(source: string, target: string): string {
  return resolveRelTarget(source.slice(0, source.lastIndexOf('/')), target);
}

/** The direct shape-tree child, including compatible-content wrappers. */
export function sourceElement(slide: SlideData, id: string): Element | undefined {
  const tree = slide.root?.child('cSld').child('spTree');
  if (!tree?.exists()) return undefined;
  return (
    tree
      .allChildren()
      .map((child) => (child.exists() ? child.element : null))
      .find(
        (element) =>
          element
            ?.getElementsByTagNameNS(
              'http://schemas.openxmlformats.org/presentationml/2006/main',
              'cNvPr',
            )[0]
            ?.getAttribute('id') === id,
      ) ?? undefined
  );
}
