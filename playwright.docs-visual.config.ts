import { defineConfig } from '@playwright/test';

export default defineConfig({
  expect: {
    timeout: 5_000,
    toHaveScreenshot: {
      animations: 'disabled',
      maxDiffPixelRatio: 0.002,
      threshold: 24 / 255,
    },
  },
  forbidOnly: Boolean(process.env.CI),
  fullyParallel: false,
  outputDir: 'test-results/playwright-docs-visual',
  projects: [
    {
      name: 'chromium',
      use: {
        browserName: 'chromium',
        deviceScaleFactor: 1,
        viewport: { height: 900, width: 1440 },
      },
    },
  ],
  reporter: [['list'], ['html', { open: 'never', outputFolder: 'playwright-report/docs-visual' }]],
  retries: 0,
  snapshotPathTemplate: '{testDir}/{testFilePath}-snapshots/{arg}-{platform}{ext}',
  testDir: './tests/docs-visual',
  testMatch: ['docs-visual.spec.ts'],
  timeout: 30_000,
  use: {
    baseURL: 'http://127.0.0.1:4176/tego-sheet/',
    colorScheme: 'light',
    locale: 'en-US',
    timezoneId: 'UTC',
    trace: 'retain-on-failure',
  },
  webServer: {
    command: 'npm run docs:build && npm run docs:serve -- --host 127.0.0.1 --port 4176',
    url: 'http://127.0.0.1:4176/tego-sheet/',
    timeout: 180_000,
    reuseExistingServer: !process.env.CI,
  },
  workers: 1,
});
