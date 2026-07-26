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
  await expect(page.getByLabel('Document JSON')).toContainText(text);
  const inspector = page.getByRole('complementary', { name: 'Playground inspector' });
  await expect(inspector.locator('li').filter({ hasText: 'onChange' }).first()).toBeVisible();
  await expect(inspector.locator('li').filter({ hasText: 'onCellEdit' }).first()).toBeVisible();
}

async function openNavigation(page: Page): Promise<void> {
  const toggle = page.getByRole('button', { name: 'Toggle navigation bar' });
  if (await toggle.isVisible()) await toggle.click();
}

async function openOutputStudio(page: Page): Promise<void> {
  await page.goto('playground?workspace=output');
  await expect(page).toHaveURL(
    new RegExp(`${PROJECT_PATH}playground\\?workspace=output&mode=uncontrolled$`),
  );
  await expect(page.getByRole('heading', { name: 'Output Studio' })).toBeVisible();
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);
}

test('project-subpath navigation loads Docs, API, Playground, and Roadmap assets without 404s', async ({
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
    ['Roadmap', 'roadmap'],
  ] as const) {
    await openNavigation(page);
    await page.getByRole('link', { name: label, exact: true }).click();
    await expect(page).toHaveURL(new RegExp(`${PROJECT_PATH}${path}/?(?:\\?.*)?$`));
    await page.waitForLoadState('networkidle');
    await page.goto('./');
  }

  expect(missingAssets).toEqual([]);
});

test('Output Studio opens directly and is linked from the printing guide', async ({ page }) => {
  await page.goto('docs/guides/printing');
  const studioLink = page.getByRole('link', { name: 'Output Studio' });
  await expect(studioLink).toHaveAttribute('href', `${PROJECT_PATH}playground?workspace=output`);
  await studioLink.click();

  await expect(page.getByRole('heading', { name: 'Output Studio' })).toBeVisible();
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);
});

test('Output Studio supports keyboard workspace switching', async ({ page }) => {
  await openOutputStudio(page);
  const outputTab = page.getByRole('tab', { name: 'Output Studio' });
  await outputTab.focus();
  await page.keyboard.press('ArrowLeft');

  await expect(page.getByRole('tab', { name: 'Spreadsheet' })).toBeFocused();
  await expect(page.getByRole('radio', { name: 'Uncontrolled' })).toBeVisible();
  await expect.poll(() => new URL(page.url()).searchParams.get('workspace')).toBe('spreadsheet');

  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('tab', { name: 'Output Studio' })).toBeFocused();
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect.poll(() => new URL(page.url()).searchParams.get('workspace')).toBe('output');
});

test('Output Studio regenerates one revision and disables stale outputs', async ({ page }) => {
  await openOutputStudio(page);
  await page.getByRole('button', { name: 'Edit template' }).click();
  await page.getByLabel('Data JSON').fill(
    JSON.stringify({
      customer: { name: 'Northwind Traders', address: 'Berlin' },
      invoice: { id: 'INV-2026-043', currency: 'EUR' },
      items: [
        { description: 'Hosting', quantity: 1, amount: 29 },
        { description: 'Support', quantity: 4, amount: 75 },
        { description: 'Training', quantity: 2, amount: 240 },
      ],
    }),
  );

  await expect(page.getByRole('status').filter({ hasText: 'Preview is stale' })).toBeVisible();
  for (const name of ['Print 2 pages', 'Download PDF', 'Download PNG page 1', 'Download XLSX']) {
    await expect(page.getByRole('button', { name })).toBeDisabled();
  }

  await page.getByRole('button', { name: 'Apply & regenerate' }).click();
  await expect(page.getByText('GeneratedDocument · revision 2')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);
  await expect(page.getByRole('button', { name: 'Download PDF' })).toBeEnabled();
});

test('Output Studio blocks invalid JSON without replacing the generated revision', async ({
  page,
}) => {
  await openOutputStudio(page);
  await page.getByRole('button', { name: 'Edit template' }).click();
  await page.getByLabel('Data JSON').fill('{');
  await page.getByRole('button', { name: 'Apply & regenerate' }).click();

  await expect(page.getByRole('alert')).toHaveText('Data must be valid JSON before regeneration.');
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Download PDF' })).toBeDisabled();
});

test('Output Studio preserves its preview when generation is blocked', async ({ page }) => {
  await openOutputStudio(page);
  await page.getByRole('button', { name: 'Edit template' }).click();
  await page.getByLabel('Expression for customer-name').fill('customer[');
  await page.getByRole('button', { name: 'Apply & regenerate' }).click();

  await expect(page.getByRole('status').filter({ hasText: 'Generation is blocked' })).toBeVisible();
  await expect(page.getByRole('list', { name: 'Blocking errors' })).toBeVisible();
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);
  await expect(page.getByRole('button', { name: 'Download PDF' })).toBeDisabled();
});

test('Output Studio zooms inside its viewport without panel overlap and keeps 44px controls', async ({
  page,
}) => {
  for (const width of [1024, 390]) {
    await page.setViewportSize({ width, height: 1100 });
    await openOutputStudio(page);

    const preview = page.getByRole('region', { name: 'Exact page preview' });
    const inputs = page.getByRole('region', { name: 'Output inputs' });
    const firstPage = page.getByRole('article', { name: /Print page/ }).first();
    const reset = page.getByRole('button', { name: 'Reset Output Studio' });
    const currentPage = page.getByLabel('Current page');
    const zoom = page.getByLabel('Preview zoom');
    const basePageBox = await firstPage.boundingBox();
    expect(basePageBox).not.toBeNull();

    await zoom.selectOption('150');
    await expect(page.getByText('Preview zoom · 150%')).toBeVisible();
    const zoomedPageBox = await firstPage.boundingBox();
    const previewBox = await preview.boundingBox();
    const inputsBox = await inputs.boundingBox();
    expect(zoomedPageBox).not.toBeNull();
    expect(previewBox).not.toBeNull();
    expect(inputsBox).not.toBeNull();
    expect(zoomedPageBox!.width).toBeGreaterThan(basePageBox!.width * 1.45);
    expect(zoomedPageBox!.height).toBeGreaterThan(basePageBox!.height * 1.45);
    expect(previewBox!.y + previewBox!.height).toBeLessThanOrEqual(inputsBox!.y);

    for (const control of [reset, currentPage, zoom]) {
      const box = await control.boundingBox();
      expect(box).not.toBeNull();
      expect(box!.height).toBeGreaterThanOrEqual(44);
    }
  }
});

test('Output Studio gives every enabled embedded workbench control a 44px target at narrow width', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 1100 });
  await openOutputStudio(page);
  await page.getByRole('button', { name: 'Edit template' }).click();

  const workbench = page.getByRole('region', { name: 'Template workbench' });
  const controls = workbench.locator(
    '.tego-sheet button:enabled, .tego-sheet select:enabled, .tego-sheet input:enabled',
  );
  expect(await controls.count()).toBeGreaterThan(0);
  const undersized = await controls.evaluateAll((elements) =>
    elements
      .map((element) => {
        const target =
          element instanceof HTMLInputElement &&
          (element.type === 'checkbox' || element.type === 'radio')
            ? (element.closest('label') ?? element)
            : element;
        const rect = target.getBoundingClientRect();
        return {
          label:
            element.getAttribute('aria-label') ??
            element.getAttribute('role') ??
            element.textContent?.trim() ??
            element.tagName,
          height: rect.height,
          width: rect.width,
        };
      })
      .filter(({ height, width }) => height < 44 || width < 44),
  );
  expect(undersized).toEqual([]);
});

test('Output Studio keeps the complete narrow preview, inputs, diagnostics, and outputs order', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 1100 });
  await openOutputStudio(page);

  const ordered = [
    page.getByRole('heading', { name: 'Exact page preview' }),
    page.getByRole('heading', { name: 'Output inputs' }),
    page.getByRole('heading', { name: 'Diagnostics' }),
    page.getByText('Print opens your system print dialog.'),
  ];
  const positions = await Promise.all(
    ordered.map(async (locator) => {
      const box = await locator.boundingBox();
      expect(box).not.toBeNull();
      return box!.y;
    }),
  );
  expect(positions).toEqual([...positions].sort((left, right) => left - right));
});

test('Output Studio exposes stubbed output actions without opening native Print', async ({
  page,
}) => {
  await openOutputStudio(page);

  const actionNames = [
    'Print 2 pages',
    'Download PDF',
    'Download PNG page 1',
    'Download XLSX',
  ] as const;
  await page.evaluate((names) => {
    const invoked: string[] = [];
    Object.defineProperty(window, '__outputStudioActions', {
      configurable: true,
      value: invoked,
    });
    for (const button of document.querySelectorAll('button')) {
      const name = button.textContent?.trim();
      if (name === undefined || !names.includes(name as (typeof names)[number])) continue;
      button.addEventListener(
        'click',
        (event) => {
          event.stopImmediatePropagation();
          invoked.push(name);
        },
        { capture: true },
      );
    }
  }, actionNames);

  for (const name of actionNames) {
    await expect(page.getByRole('button', { name })).toBeEnabled();
    await page.getByRole('button', { name }).click();
  }

  await expect
    .poll(() =>
      page.evaluate(
        () =>
          (
            window as typeof window & {
              __outputStudioActions: string[];
            }
          ).__outputStudioActions,
      ),
    )
    .toEqual(actionNames);
});

test('Roadmap exposes five dependency phases and two non-interactive planned items', async ({
  page,
}) => {
  await page.goto('roadmap');

  await expect(page.getByRole('heading', { level: 1, name: 'Product roadmap' })).toBeVisible();
  await expect(page.locator('[data-roadmap-phase]')).toHaveCount(5);
  await expect(page.locator('[data-roadmap-item]')).toHaveCount(2);
  await expect(page.getByText('Planned', { exact: true })).toHaveCount(2);
  await expect(page.getByRole('checkbox')).toHaveCount(0);
  await expect(page.getByRole('link', { name: 'XLSX template output' })).toHaveAttribute(
    'href',
    `${PROJECT_PATH}docs/roadmap/template-printing`,
  );
  await expect(
    page.getByRole('link', { name: 'CSV/TSV, XLSX and ODS interchange' }),
  ).toHaveAttribute('href', `${PROJECT_PATH}docs/roadmap/formulas-data`);

  const viewport = page.viewportSize();
  expect(viewport).not.toBeNull();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(viewport?.width);
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
  await expect(page.getByLabel('Document JSON')).not.toContainText('Uncontrolled browser edit');
  await expect(page.getByText('Interact with the sheet to inspect callbacks.')).toBeVisible();

  await page.getByRole('radio', { name: 'Controlled', exact: true }).check();
  await waitForSheet(page, 'controlled');
  await editCell(page, 'Controlled browser edit');
  await page.reload();
  await waitForSheet(page, 'controlled');
  await expect(page.getByLabel('Document JSON')).not.toContainText('Controlled browser edit');
  await expect(page.getByLabel('Document JSON')).toContainText('Keyboard');
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
