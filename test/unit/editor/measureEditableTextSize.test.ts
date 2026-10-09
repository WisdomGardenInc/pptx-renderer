import { afterEach, describe, expect, it, vi } from 'vitest';
import { measureEditableTextSize } from '../../../src/editor/measureEditableTextSize';

afterEach(() => {
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

function input() {
  const root = document.createElement('div');
  Object.assign(root.style, {
    width: 'max-content',
    height: 'auto',
    minWidth: '400px',
    minHeight: '80px',
  });
  document.body.append(root);
  vi.spyOn(root, 'offsetWidth', 'get').mockReturnValue(460);
  vi.spyOn(root, 'offsetHeight', 'get').mockReturnValue(100);
  return root;
}

const shape = { width: 600, height: 120 };
const text = { width: 400, height: 80 };

describe('measureEditableTextSize', () => {
  it('preserves the difference between shape extents and inset text bounds without compounding', () => {
    const root = input();
    const grown = measureEditableTextSize(root, shape, shape, text, false);
    expect(grown).toEqual({ width: 660, height: 140 });
    expect(measureEditableTextSize(root, grown, shape, text, false)).toEqual(grown);
  });

  it('releases both authored minima after a text reduction', () => {
    const root = input();
    const current = { width: 800, height: 300 };
    expect(measureEditableTextSize(root, current, shape, text, true)).toEqual({
      width: 660,
      height: 140,
    });
    expect(root.style.minWidth).toBe('0px');
    expect(root.style.minHeight).toBe('0px');
  });

  it('retains dimensions while disconnected or not measurable', () => {
    const root = input();
    root.remove();
    expect(measureEditableTextSize(root, shape, shape, text, false)).toEqual(shape);
    document.body.append(root);
    vi.spyOn(root, 'offsetWidth', 'get').mockReturnValue(0);
    vi.spyOn(root, 'offsetHeight', 'get').mockReturnValue(0);
    expect(measureEditableTextSize(root, shape, shape, text, true)).toEqual(shape);
  });
});
