import { defineConfig } from '@playwright/test';

// One server, one browser, one worker: the editor E2E is small and the dev machine is not.
export default defineConfig({
  testDir: '.',
  testMatch: '*.spec.ts',
  globalSetup: './server.ts',
  workers: 1,
  fullyParallel: false,
  retries: 0,
  timeout: 60_000,
  reporter: process.env.CI ? [['list'], ['github']] : 'list',
  use: {
    browserName: 'chromium',
    headless: true,
    viewport: { width: 1400, height: 900 },
    acceptDownloads: true,
    trace: 'retain-on-failure',
  },
});
