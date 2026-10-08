import { describe, expect, it } from 'vitest';
import { waitForSlideImages } from '../../../src/renderer/slideImages';

describe('slide image readiness', () => {
  it('waits for decoding after load and settles image failures', async () => {
    const root = document.createElement('div');
    const image = document.createElement('img');
    image.src = '/picture.png';
    root.appendChild(image);
    let decoded!: () => void;
    image.decode = () =>
      new Promise((resolve) => {
        decoded = resolve;
      });
    let ready = false;
    const waiting = waitForSlideImages(root, new AbortController().signal).then(() => {
      ready = true;
    });
    image.dispatchEvent(new Event('load'));
    await Promise.resolve();
    expect(ready).toBe(false);
    decoded();
    await waiting;
    expect(ready).toBe(true);
    const failing = waitForSlideImages(root, new AbortController().signal);
    image.dispatchEvent(new Event('error'));
    await failing;
  });

  it('aborts waits on SVG and CSS background image resources', async () => {
    const root = document.createElement('div');
    root.style.backgroundImage = 'url("/background.png")';
    root.innerHTML = '<svg><image href="/svg-image.png"/></svg>';
    const controller = new AbortController();
    let ready = false;
    const waiting = waitForSlideImages(root, controller.signal).then(() => {
      ready = true;
    });
    await Promise.resolve();
    expect(ready).toBe(false);
    controller.abort();
    await waiting;
    expect(ready).toBe(true);
  });
});
