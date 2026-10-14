import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './test',
  outputDir: '../../docs/agent-tmp/editor-tests',
  use: {
    baseURL: 'http://127.0.0.1:5274',
    browserName: 'chromium',
    headless: true,
  },
  webServer: {
    command: 'pnpm dev --host 127.0.0.1 --port 5274 --strictPort',
    cwd: '.',
    url: 'http://127.0.0.1:5274',
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
});
