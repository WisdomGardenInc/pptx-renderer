import { defineConfig } from 'vite';
import { resolve } from 'path';

/**
 * Builds the `./editor` subpath entry.
 *
 * Separate from the main config because the editor core lives under `apps/` and must
 * not be bundled into the library: consumers import both, and a second copy of the
 * renderer inside the editor bundle would mean two module registries, two media
 * caches, and `instanceof` checks that quietly stop matching.
 *
 * Shipped as JS rather than the TypeScript source a bundler could transpile itself —
 * Turbopack will not resolve a package export that points at a `.ts` file.
 */
export default defineConfig({
  define: {
    'process.env.NODE_ENV': '"production"',
  },
  resolve: {
    alias: {
      '@wisdomgarden/pptx-renderer': resolve(__dirname, 'src/index.ts'),
    },
  },
  build: {
    emptyOutDir: false,
    lib: {
      entry: resolve(__dirname, 'apps/editor/src/editor/index.ts'),
      formats: ['es'],
      fileName: () => 'editor.es.js',
    },
    rollupOptions: {
      // The renderer is resolved through the alias above at type-check time but stays
      // external at runtime, so the host loads exactly one copy of it.
      external: (id) =>
        id === '@wisdomgarden/pptx-renderer' ||
        id === 'echarts' ||
        id.startsWith('echarts/') ||
        id === 'jszip' ||
        id === 'pdfjs-dist',
    },
  },
});
