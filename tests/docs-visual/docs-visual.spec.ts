import { expect, test, type Locator, type Page } from '@playwright/test';

const DESKTOP = { height: 900, width: 1440 } as const;
const NARROW = { height: 844, width: 390 } as const;

test.beforeEach(async ({ page }) => {
  await page.emulateMedia({ colorScheme: 'light', reducedMotion: 'reduce' });
});

async function waitForFonts(page: Page): Promise<void> {
  const status = await page.evaluate(async () => {
    await document.fonts.ready;
    return document.fonts.status;
  });
  expect(status).toBe('loaded');
}

async function waitForSheet(page: Page, mode: 'controlled' | 'uncontrolled'): Promise<void> {
  await expect(page.locator('[data-tego-sheet]')).toHaveAttribute('data-mode', mode);
  const canvas = page.locator('.tego-sheet__canvas');
  await expect(canvas).toBeVisible();
  await expect
    .poll(async () => {
      const box = await canvas.boundingBox();
      return box !== null && box.width > 0 && box.height > 0;
    })
    .toBe(true);
}

function volatileInspectorMasks(page: Page): Locator[] {
  return [
    page.locator('[aria-label^="Event "] > strong'),
    page.locator('[aria-label^="Event "] time'),
  ];
}

test('home desktop', async ({ page }) => {
  await page.setViewportSize(DESKTOP);
  await page.goto('./');
  await waitForSheet(page, 'controlled');
  await waitForFonts(page);
  await expect(page).toHaveScreenshot('home-desktop.png', {
    animations: 'disabled',
    fullPage: true,
  });
});

test('Quick Start desktop', async ({ page }) => {
  await page.setViewportSize(DESKTOP);
  await page.goto('docs/getting-started/quick-start');
  await expect(page.getByRole('heading', { level: 1 })).toHaveText('Quick Start');
  await waitForFonts(page);
  await expect(page).toHaveScreenshot('quick-start-desktop.png', {
    animations: 'disabled',
    fullPage: true,
  });
});

test('Roadmap desktop', async ({ page }) => {
  await page.setViewportSize(DESKTOP);
  await page.goto('roadmap');
  await expect(page.getByRole('heading', { level: 1, name: 'Product roadmap' })).toBeVisible();
  await waitForFonts(page);
  await expect(page).toHaveScreenshot('roadmap-desktop.png', {
    animations: 'disabled',
    fullPage: true,
  });
});

test('Playground desktop Controlled', async ({ page }) => {
  await page.setViewportSize(DESKTOP);
  await page.goto('playground?mode=controlled');
  await waitForSheet(page, 'controlled');
  await waitForFonts(page);
  await expect(page).toHaveScreenshot('playground-controlled-desktop.png', {
    animations: 'disabled',
    mask: volatileInspectorMasks(page),
  });
});

test('Playground narrow Uncontrolled', async ({ page }) => {
  await page.setViewportSize(NARROW);
  await page.goto('playground?mode=uncontrolled');
  await waitForSheet(page, 'uncontrolled');
  await waitForFonts(page);
  await expect(page).toHaveScreenshot('playground-uncontrolled-narrow.png', {
    animations: 'disabled',
    fullPage: true,
    mask: volatileInspectorMasks(page),
  });
});
