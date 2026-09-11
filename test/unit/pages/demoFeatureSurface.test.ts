import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { resolve, dirname } from 'node:path';
import { describe, expect, it } from 'vitest';

const __dirname = dirname(fileURLToPath(import.meta.url));
const demoPath = resolve(__dirname, '../../../demo/index.html');
const devPagePath = resolve(__dirname, '../../pages/index.html');
const demoViteConfigPath = resolve(__dirname, '../../../vite.config.demo.ts');
const docsExampleSamplePath = resolve(
  __dirname,
  '../../../docs/example/1-chart-and-complex/source.pptx',
);
const duplicatePublicSamplePath = resolve(
  __dirname,
  '../../../demo/public/samples/chart-and-complex.pptx',
);
const demoHtml = readFileSync(demoPath, 'utf-8');
const devHtml = readFileSync(devPagePath, 'utf-8');
const demoViteConfig = readFileSync(demoViteConfigPath, 'utf-8');

const pages = [
  ['public demo', demoHtml],
  ['dev preview page', devHtml],
] as const;

describe('public demo feature surface', () => {
  it('exposes model text search controls and viewer search API usage', () => {
    for (const [label, html] of pages) {
      expect(html, label).toContain('id="search-input"');
      expect(html, label).toContain('viewer.searchText');
      expect(html, label).toContain('search-results');
      expect(html, label).toContain('id="search-prev"');
      expect(html, label).toContain('id="search-next"');
      expect(html, label).toContain('id="search-position"');
      expect(html, label).toContain('id="search-match-case"');
      expect(html, label).toContain('title="Match case"');
      expect(html, label).toContain('matchCase: searchMatchCase.checked');
      expect(html, label).toContain("searchMatchCase.addEventListener('change', runSearch)");
      expect(html, label).toContain('activateSearchResult');
      expect(html, label).toContain('renderHighlightedSnippet');
      expect(html, label).toContain('viewer.highlightSearchResult');
      expect(html, label).toContain('viewer.clearSearchHighlights');
    }
  });

  it('exposes lazy scaled slide thumbnail previews', () => {
    for (const [label, html] of pages) {
      expect(html, label).toContain('thumbnail-list');
      expect(html, label).toContain('IntersectionObserver');
      expect(html, label).toContain('renderThumbnailToContainer');
      expect(html, label).toContain('cleanupThumbnails');
      expect(html, label).toContain('width: 96px');
      expect(html, label).toContain('height: 54px');
      expect(html, label).toContain('width: 112px');
      expect(html, label).toContain('min-width: 112px');
      expect(html, label).toContain('max-width: 112px');
      expect(html, label).toContain('box-sizing: border-box');
      expect(html, label).toContain('padding: 6px');
      expect(html, label).toContain('{ width: 96 }');
    }
  });

  it('keeps thumbnail selection highlight layout-stable', () => {
    for (const [label, html] of pages) {
      expect(html, label).toContain('box-shadow: inset 0 0 0 2px var(--accent)');
      expect(html, label).not.toContain(
        '.search-result:hover,\n      .search-result.active,\n      .thumbnail-item.active',
      );
      expect(html, label).not.toContain('.thumbnail-item.active {\n        border-color');
      expect(html, label).not.toContain('box-shadow 120ms ease');
    }
  });

  it('uses the viewer node-level search highlight API without renderer text rewriting', () => {
    for (const [label, html] of pages) {
      expect(html, label).toContain('search-highlight-overlay');
      expect(html, label).not.toContain('function showSearchHighlight');
      expect(html, label).not.toContain('function clearSearchHighlight');
      expect(html, label).not.toContain('textContent.replace');
    }
  });

  it('offers primary empty-state actions in the public demo', () => {
    expect(demoHtml).toContain('id="empty-upload-btn"');
    expect(demoHtml).toContain('id="empty-sample-btn"');
    expect(demoHtml).toContain('emptyUploadBtn?.addEventListener');
    expect(demoHtml).toContain('emptySampleBtn?.addEventListener');
    expect(demoHtml).toContain('fileInput.click()');
    expect(demoHtml).toContain("loadSample('samples/chart-and-complex.pptx')");
  });

  it('keeps the public demo in list mode instead of exposing a broken mode toggle', () => {
    expect(demoHtml).not.toContain('id="mode-toggle"');
    expect(demoHtml).not.toContain('data-mode="slide"');
    expect(demoHtml).not.toContain('function switchMode');
    expect(demoHtml).not.toContain('currentMode');
    expect(demoHtml).toContain("renderMode: 'list'");
  });

  it('uses large-deck viewer options in the public demo', () => {
    expect(demoHtml).toContain('RECOMMENDED_ZIP_LIMITS');
    expect(demoHtml).toContain('function createViewerOpenOptions()');
    expect(demoHtml).toContain('zipLimits: RECOMMENDED_ZIP_LIMITS');
    expect(demoHtml).toContain('lazySlides: true');
    expect(demoHtml).toContain('lazyMedia: true');
    expect(demoHtml).toContain('listOptions: demoListOptions');
    expect(demoHtml).toContain('PptxViewer.open(buffer, container, createViewerOpenOptions())');
  });

  it('serves the public demo sample from the docs example source deck', () => {
    expect(existsSync(docsExampleSamplePath)).toBe(true);
    expect(existsSync(duplicatePublicSamplePath)).toBe(false);
    expect(demoHtml).toContain('data-file="samples/chart-and-complex.pptx"');
    expect(demoViteConfig).toContain('docs/example/1-chart-and-complex/source.pptx');
    expect(demoViteConfig).toContain('fileName: samplePublicFile');
    expect(demoViteConfig).toContain('publicDir: false');
  });

  it('lets the public demo render through a published renderer build', () => {
    expect(demoHtml).toContain('id="version-select"');
    expect(demoHtml).toContain("versionSelect.addEventListener('change'");
    expect(demoHtml).toContain('switchRendererVersion');

    // Both npm scopes must be queried: ./browser exists from 1.2.4 (old scope) on.
    expect(demoHtml).toContain("'@wisdomgarden/pptx-renderer', '@aiden0z/pptx-renderer'");
    expect(demoHtml).toContain('registry.npmjs.org/${pkg}');
    expect(demoHtml).toContain("exports?.['./browser']");
    expect(demoHtml).toContain('.filter((build) => build.entry)');
  });

  it('rebinds the renderer entry points instead of pinning the bundled import', () => {
    expect(demoHtml).toContain("import * as sourceLib from '../src/index.ts'");
    expect(demoHtml).toContain('let PptxViewer = sourceLib.PptxViewer');
    expect(demoHtml).toContain('let RECOMMENDED_ZIP_LIMITS = sourceLib.RECOMMENDED_ZIP_LIMITS');
    expect(demoHtml).toContain('PptxViewer = mod.PptxViewer');

    // The dev page always renders the working tree, so it keeps the plain import.
    expect(devHtml).not.toContain('version-select');
  });

  it('resolves the published entry path from the registry manifest', () => {
    // Hardcoding the artifact basename would break on a future build-output rename.
    expect(demoHtml).toContain('function browserEntryPath(meta)');
    expect(demoHtml).toContain('entry?.import ?? entry?.default');
    expect(demoHtml).toContain('cdn.jsdelivr.net/npm/${build.pkg}@${build.version}/${build.entry}');
  });

  it('re-renders the loaded deck when the renderer version changes', () => {
    expect(demoHtml).toContain('cachedLabel');
    expect(demoHtml).toContain('cachedBytes');
    // Trailing arguments are open on purpose: what matters is that a version
    // switch re-renders the cached deck with its cached metadata, not how many
    // extras (slide position, and so on) renderBuffer grows.
    expect(demoHtml).toMatch(
      /if \(cachedBuffer\) await renderBuffer\(\s*cachedBuffer,\s*cachedLabel,\s*cachedBytes\s*[,)]/,
    );
  });

  it('caches loaded renderer builds and reverts a failed version switch', () => {
    expect(demoHtml).toContain('loadedRenderers');
    expect(demoHtml).toContain('loadedRenderers.set(value, mod)');
    expect(demoHtml).toContain('versionSelect.value = previous');
    expect(demoHtml).toContain('Failed to load renderer ${value}');
    expect(demoHtml).toContain('versionSelect.disabled = false');
  });

  it('keeps the public demo empty state visually anchored', () => {
    const previewCss = demoHtml.match(/\.empty-slide-preview\s*\{([^}]*)\}/)?.[1] ?? '';

    expect(demoHtml).toContain('empty-slide-preview');
    expect(demoHtml).toContain('empty-feature-icon');
    expect(demoHtml).toContain('Open a deck or launch the sample');
    expect(previewCss).toContain('width: 100%;');
  });
});
