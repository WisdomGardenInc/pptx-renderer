const FILL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill']);

/** Find the effective ancestor group paint for an element using a:grpFill. */
export function inheritedGroupFill(element: Element): Element | undefined {
  for (let parent = element.parentElement; parent; parent = parent.parentElement) {
    if (parent.localName !== 'grpSp') continue;
    const properties = Array.from(parent.children).find((child) => child.localName === 'grpSpPr');
    const fill =
      properties &&
      Array.from(properties.children).find((child) => FILL_NAMES.has(child.localName));
    if (fill) return fill;
  }
  return undefined;
}
