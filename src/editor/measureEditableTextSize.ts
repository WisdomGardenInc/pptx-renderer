export interface TextSize {
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
  const delta = measured - text;
  const fitted = shape + (Math.abs(delta) < 1 / 64 ? 0 : Math.ceil(delta));
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

export function measureTextContentSize(
  root: HTMLElement,
  shape: TextSize,
  text: TextSize,
  wrap: boolean,
): TextSize {
  if (!root.isConnected || root.style.writingMode)
    return { width: shape.width, height: shape.height };
  const copy = root.cloneNode(true) as HTMLElement;
  copy.removeAttribute('contenteditable');
  copy.removeAttribute('id');
  delete copy.dataset.editing;
  const computed = root.ownerDocument.defaultView?.getComputedStyle(root);
  Object.assign(copy.style, {
    position: 'fixed',
    left: '-100000px',
    top: '0',
    width: wrap ? `${text.width}px` : 'max-content',
    height: 'auto',
    minWidth: '0px',
    minHeight: '0px',
    maxWidth: 'none',
    maxHeight: 'none',
    transform: 'none',
    justifyContent: 'flex-start',
    whiteSpace: wrap ? 'normal' : 'nowrap',
    overflow: 'visible',
    fontFamily: computed?.fontFamily ?? '',
    visibility: 'hidden',
    pointerEvents: 'none',
  });
  root.ownerDocument.body.appendChild(copy);
  try {
    return measureEditableTextSize(copy, shape, shape, text, true);
  } finally {
    copy.remove();
  }
}
