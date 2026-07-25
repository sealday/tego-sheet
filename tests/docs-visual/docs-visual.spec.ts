import { expect, test, type Locator, type Page } from '@playwright/test';

const DESKTOP = { height: 900, width: 1440 } as const;
const OUTPUT_DESKTOP = { height: 1100, width: 1440 } as const;
const OUTPUT_INTERMEDIATE = { height: 1100, width: 1024 } as const;
const OUTPUT_NARROW = { height: 1000, width: 390 } as const;
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

async function openOutputStudio(
  page: Page,
  viewport: { readonly height: number; readonly width: number },
): Promise<void> {
  await page.setViewportSize(viewport);
  await page.goto('playground?workspace=output');
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);
  await waitForFonts(page);
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

for (const [name, viewport, snapshot] of [
  ['desktop', OUTPUT_DESKTOP, 'output-studio-ready-desktop.png'],
  ['intermediate', OUTPUT_INTERMEDIATE, 'output-studio-ready-intermediate.png'],
  ['narrow', OUTPUT_NARROW, 'output-studio-ready-narrow.png'],
] as const) {
  test(`Output Studio ${name} ready`, async ({ page }) => {
    await openOutputStudio(page, viewport);
    await expect(page).toHaveScreenshot(snapshot, { animations: 'disabled' });
  });
}

test('Output Studio desktop stale', async ({ page }) => {
  await openOutputStudio(page, OUTPUT_DESKTOP);
  await page.getByRole('button', { name: 'Edit template' }).click();
  await page.getByLabel('Data JSON').fill(
    JSON.stringify({
      customer: { name: 'Northwind Traders', address: 'Berlin' },
      invoice: { id: 'INV-2026-043', currency: 'EUR' },
      items: [{ description: 'Hosting', quantity: 1, amount: 29 }],
    }),
  );
  await expect(page.getByRole('status').filter({ hasText: 'Preview is stale' })).toBeVisible();
  await expect(page).toHaveScreenshot('output-studio-stale-desktop.png', {
    animations: 'disabled',
  });
});

test('Output Studio desktop blocked diagnostic', async ({ page }) => {
  await openOutputStudio(page, OUTPUT_DESKTOP);
  await page.getByRole('button', { name: 'Edit template' }).click();
  await page.getByLabel('Expression for customer-name').fill('customer[');
  await page.getByRole('button', { name: 'Apply & regenerate' }).click();
  await expect(page.getByRole('status').filter({ hasText: 'Generation is blocked' })).toBeVisible();
  await expect(page.getByRole('list', { name: 'Generation diagnostics' })).toBeVisible();
  await page.evaluate(() => window.scrollTo(0, 0));
  await expect(page).toHaveScreenshot('output-studio-blocked-desktop.png', {
    animations: 'disabled',
    maxDiffPixelRatio: 0,
  });
});
