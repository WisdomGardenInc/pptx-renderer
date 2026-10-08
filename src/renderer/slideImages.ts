/** Wait for HTML, SVG and CSS images to load and decode; errors and aborts settle.
 * Call after asynchronous rendering has populated the slide DOM. */
function waitForImage(image: HTMLImageElement, signal: AbortSignal): Promise<void> {
  return new Promise((resolve) => {
    const finish = () => {
      image.removeEventListener('load', loaded);
      image.removeEventListener('error', finish);
      signal.removeEventListener('abort', finish);
      resolve();
    };
    const loaded = () => {
      if (signal.aborted || typeof image.decode !== 'function') {
        finish();
        return;
      }
      // A load event alone does not mean the browser has decoded a paintable frame.
      void image.decode().then(finish, finish);
    };
    if (signal.aborted) {
      resolve();
      return;
    }
    image.addEventListener('load', loaded, { once: true });
    image.addEventListener('error', finish, { once: true });
    signal.addEventListener('abort', finish, { once: true });
    if (image.complete) {
      if (image.naturalWidth > 0) loaded();
      else finish();
    }
  });
}

export async function waitForSlideImages(root: HTMLElement, signal: AbortSignal): Promise<void> {
  const images = Array.from(root.querySelectorAll('img'));
  const urls = new Set<string>();
  // PPTX pictures can also render as SVG images or CSS background layers.
  for (const element of [root, ...Array.from(root.querySelectorAll('*'))]) {
    if (element.localName === 'image') {
      const href =
        element.getAttribute('href') ??
        element.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
      if (href) urls.add(href);
    }
    const background = (element as HTMLElement | SVGElement).style?.backgroundImage ?? '';
    for (const match of background.matchAll(/url\(\s*(?:"([^"]*)"|'([^']*)'|([^)]*))\s*\)/g)) {
      const url = (match[1] ?? match[2] ?? match[3]).trim();
      if (url) urls.add(url);
    }
  }
  for (const url of urls) {
    const image = root.ownerDocument.createElement('img');
    image.src = url;
    images.push(image);
  }
  await Promise.all(images.map((image) => waitForImage(image, signal)));
}
