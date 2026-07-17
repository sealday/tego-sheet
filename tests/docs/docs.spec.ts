import { expect, test, type Page, type Response } from '@playwright/test';

const PROJECT_PATH = '/tego-sheet/';
const MODES = [
  ['Uncontrolled', 'uncontrolled'],
  ['Controlled', 'controlled'],
  ['Custom Chrome', 'custom-chrome'],
  ['Locales', 'locales'],
  ['Legacy JSON', 'legacy-json'],
] as const;

async function waitForSheet(page: Page, mode: string): Promise<void> {
  await expect(page.locator('[data-tego-sheet]')).toHaveAttribute('data-mode', mode);
  await expect(page.locator('.tego-sheet__canvas')).toBeVisible();
}

test.beforeEach(async ({ page }) => {
  await page.emulateMedia({ colorScheme: 'light', reducedMotion: 'reduce' });
});

async function openPlayground(page: Page, mode = 'uncontrolled'): Promise<void> {
  await page.goto(`playground?mode=${mode}`);
  await waitForSheet(page, mode);
}

async function cellPoint(page: Page, row: number, column: number) {
  const canvas = page.locator('.tego-sheet__canvas');
  const box = await canvas.boundingBox();
  if (box === null) throw new Error('Canvas has no box');
  const clientSize = await canvas.evaluate((element) => ({
    height: element.clientHeight,
    width: element.clientWidth,
  }));
  const scaleX = clientSize.width > 0 ? box.width / clientSize.width : 1;
  const scaleY = clientSize.height > 0 ? box.height / clientSize.height : 1;
  return {
    x: box.x + (60 + column * 100 + 50) * scaleX,
    y: box.y + (25 + row * 25 + 12.5) * scaleY,
  };
}

async function editCell(page: Page, text: string): Promise<void> {
  const point = await cellPoint(page, 1, 0);
  await page.mouse.click(point.x, point.y);
  await page.keyboard.press('F2');
  const editor = page.getByRole('textbox', { name: 'Cell editor' });
  await expect(editor).toBeFocused();
  await editor.fill(text);
  await page.keyboard.press('Enter');
  await expect(page.getByLabel('Workbook JSON')).toContainText(text);
  const inspector = page.getByRole('complementary', { name: 'Playground inspector' });
  await expect(inspector.locator('li').filter({ hasText: 'onChange' }).first()).toBeVisible();
  await expect(inspector.locator('li').filter({ hasText: 'onCellEdit' }).first()).toBeVisible();
}

async function openNavigation(page: Page): Promise<void> {
  const toggle = page.getByRole('button', { name: 'Toggle navigation bar' });
  if (await toggle.isVisible()) await toggle.click();
}

test('project-subpath navigation loads Docs, API, and Playground assets without 404s', async ({
  page,
}) => {
  const missingAssets: string[] = [];
  const inspectResponse = (response: Response): void => {
    const type = response.request().resourceType();
    if ((type === 'script' || type === 'stylesheet') && response.status() === 404)
      missingAssets.push(response.url());
  };
  page.on('response', inspectResponse);

  await page.goto('./');
  await expect(page).toHaveURL(new RegExp(`${PROJECT_PATH.replaceAll('/', '\\/')}$`));
  await expect(page.getByRole('heading', { level: 1 })).toContainText(
    'Spreadsheet UI that belongs in your React application.',
  );

  for (const [label, path] of [
    ['Docs', 'docs/getting-started/installation'],
    ['API', 'docs/api'],
    ['Playground', 'playground'],
  ] as const) {
    await openNavigation(page);
    await page.getByRole('link', { name: label, exact: true }).click();
    await expect(page).toHaveURL(new RegExp(`${PROJECT_PATH}${path}/?(?:\\?.*)?$`));
    await page.waitForLoadState('networkidle');
    await page.goto('./');
  }

  expect(missingAssets).toEqual([]);
});

test('all five public presets follow URL history and reload behavior', async ({ page }) => {
  await openPlayground(page);

  for (const [label, mode] of MODES) {
    await page.getByRole('radio', { name: label, exact: true }).check();
    await expect.poll(() => new URL(page.url()).searchParams.get('mode')).toBe(mode);
    await waitForSheet(page, mode === 'controlled' ? 'controlled' : 'uncontrolled');
  }

  await page.goBack();
  await expect(page.getByRole('radio', { name: 'Locales', exact: true })).toBeChecked();
  await expect.poll(() => new URL(page.url()).searchParams.get('mode')).toBe('locales');
  await page.goBack();
  await expect(page.getByRole('radio', { name: 'Custom Chrome', exact: true })).toBeChecked();
  await page.goForward();
  await expect(page.getByRole('radio', { name: 'Locales', exact: true })).toBeChecked();

  await page.reload();
  await expect(page.getByRole('radio', { name: 'Locales', exact: true })).toBeChecked();
  await waitForSheet(page, 'uncontrolled');
});

test('real Canvas edits update and reset uncontrolled and controlled public inspectors', async ({
  page,
}) => {
  await openPlayground(page);
  await editCell(page, 'Uncontrolled browser edit');

  await page.getByRole('button', { name: 'Reset mode' }).click();
  await expect(page.getByLabel('Workbook JSON')).not.toContainText('Uncontrolled browser edit');
  await expect(page.getByText('Interact with the sheet to inspect callbacks.')).toBeVisible();

  await page.getByRole('radio', { name: 'Controlled', exact: true }).check();
  await waitForSheet(page, 'controlled');
  await editCell(page, 'Controlled browser edit');
  await page.reload();
  await waitForSheet(page, 'controlled');
  await expect(page.getByLabel('Workbook JSON')).not.toContainText('Controlled browser edit');
  await expect(page.getByLabel('Workbook JSON')).toContainText('Keyboard');
});

test('narrow consumers stack the inspector below the spreadsheet', async ({ page }) => {
  await page.setViewportSize({ height: 844, width: 390 });
  await openPlayground(page);
  const sheetPanel = page.locator('[data-testid="preset-boundary"] > div').first();
  const inspector = page.getByRole('complementary', { name: 'Playground inspector' });
  await expect
    .poll(async () => {
      const [sheetBox, inspectorBox] = await Promise.all([
        sheetPanel.boundingBox(),
        inspector.boundingBox(),
      ]);
      if (sheetBox === null || inspectorBox === null) return false;
      return inspectorBox.y >= sheetBox.y + sheetBox.height;
    })
    .toBe(true);
});
