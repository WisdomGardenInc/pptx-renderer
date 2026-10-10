import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath, URL } from 'node:url';

// The editor consumes the pptx-renderer library directly from source (not dist) so it
// always picks up the latest renderer, including the `onNodeRendered` editing hook.
// The library's own runtime deps (jszip, echarts, ...) resolve from the repo-root
// node_modules, which Vite finds by walking up from the aliased source files.
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@wisdomgarden/pptx-renderer': fileURLToPath(
        new URL('../../src/index.ts', import.meta.url),
      ),
    },
  },
  server: {
    port: 5273,
    // Allow serving library source and sample decks from the repo root during dev.
    fs: {
      allow: [fileURLToPath(new URL('../../', import.meta.url))],
    },
  },
});
