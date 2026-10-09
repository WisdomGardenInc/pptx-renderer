interface TextSize {
  width: number;
  height: number;
}

function fittedDimension(
  root: HTMLElement,
  axis: 'width' | 'height',
  computed: CSSStyleDeclaration | undefined,
  current: number,
  shape: number,
  text: number,
  fitContent: boolean,
): number {
  const computedSize = Number.parseFloat(computed?.[axis] ?? '');
  const measured = Number.isFinite(computedSize)
    ? computedSize
    : root[axis === 'width' ? 'offsetWidth' : 'offsetHeight'];
  if (measured <= 0) return current;
  const fitted = shape + Math.ceil(measured - text);
  return fitContent ? Math.max(1, fitted) : Math.max(current, fitted);
}

export function measureEditableTextSize(
  root: HTMLElement,
  current: TextSize,
  shape: TextSize,
  text: TextSize,
  fitContent: boolean,
): TextSize {
  const fitHeight = root.style.height === 'auto';
  const fitWidth = fitHeight && root.style.width === 'max-content';
  if (fitContent) {
    if (fitHeight && root.style.minHeight !== '0px') root.style.minHeight = '0px';
    if (fitWidth && root.style.minWidth !== '0px') root.style.minWidth = '0px';
  }
  if (!root.isConnected) return { width: current.width, height: current.height };
  const computed = root.ownerDocument.defaultView?.getComputedStyle(root);
  return {
    width: fitWidth
      ? fittedDimension(root, 'width', computed, current.width, shape.width, text.width, fitContent)
      : current.width,
    height: fitHeight
      ? fittedDimension(
          root,
          'height',
          computed,
          current.height,
          shape.height,
          text.height,
          fitContent,
        )
      : current.height,
  };
}
