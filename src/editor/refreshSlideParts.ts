import { materializeSlideNodes, type PresentationData } from '../model/Presentation';
import type { PptxFiles } from '../parser/ZipParser';
import { parseRels } from '../parser/RelParser';
import { relationshipsPart, resolvePartPath } from './parts';

/** Refresh changed parts without rebuilding shared themes, masters or untouched slides.
 * Package structure changes still require buildPresentation(). */
export function refreshSlideParts(
  presentation: PresentationData,
  files: PptxFiles,
  indices: readonly number[],
): PresentationData {
  const slides = [...presentation.slides];
  const slideToLayout = new Map(presentation.slideToLayout);
  for (const index of new Set(indices)) {
    const previous = slides[index];
    if (!previous) throw new RangeError(`Unknown slide index ${index}`);
    const xml = files.slides.get(previous.slidePath);
    if (xml == null) throw new Error(`Missing slide part ${previous.slidePath}`);
    const rels = parseRels(
      files.slideRels.get(relationshipsPart(previous.slidePath)) ?? '<Relationships/>',
    );
    const layout = Array.from(rels.values()).find((rel) => rel.type.endsWith('/slideLayout'));
    const layoutIndex = layout ? resolvePartPath(previous.slidePath, layout.target) : '';
    if (layoutIndex) slideToLayout.set(index, layoutIndex);
    else slideToLayout.delete(index);
    slides[index] = {
      ...previous,
      rels,
      layoutIndex,
      nodes: [],
      root: undefined,
      sourceXml: xml,
      nodesMaterialized: false,
      placeholderInheritanceResolved: false,
    };
  }
  const updated = { ...presentation, slides, slideToLayout };
  for (const index of new Set(indices)) materializeSlideNodes(updated, slides[index]);
  return updated;
}
