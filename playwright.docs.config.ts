import { defineConfig } from '@playwright/test';

export default defineConfig({
  expect: {
    timeout: 5_000,
  },
  forbidOnly: Boolean(process.env.CI),
  fullyParallel: false,
  outputDir: 'test-results/playwright-docs',
  projects: [
    {
      name: 'chromium-desktop',
      use: { browserName: 'chromium', viewport: { height: 900, width: 1440 } },
    },
    {
      name: 'chromium-narrow-touch',
      use: {
        browserName: 'chromium',
        hasTouch: true,
        isMobile: true,
        viewport: { height: 844, width: 390 },
      },
    },
  ],
  reporter: [['list'], ['html', { open: 'never', outputFolder: 'playwright-report/docs' }]],
  retries: 0,
  testDir: './tests/docs',
  testMatch: ['docs.spec.ts'],
  timeout: 30_000,
  use: {
    baseURL: 'http://127.0.0.1:4175/tego-sheet/',
    colorScheme: 'light',
    locale: 'en-US',
    timezoneId: 'UTC',
    trace: 'retain-on-failure',
  },
  webServer: {
    command: 'npm run docs:build && npm run docs:serve -- --host 127.0.0.1 --port 4175',
    url: 'http://127.0.0.1:4175/tego-sheet/',
    timeout: 180_000,
    reuseExistingServer: !process.env.CI,
  },
  workers: 1,
});
