import { expect, test, type Page } from '@playwright/test';
import { visualFixtures } from './fixtures';
import { namedMasks } from './masks';
import { geometryParityToken, printableCellsParityToken, visualParityByFixture } from './parity';

interface Rect {
  readonly height: number;
  readonly width: number;
  readonly x: number;
  readonly y: number;
}

async function openFixture(page: Page, name: string): Promise<void> {
  await page.emulateMedia({ colorScheme: 'light', reducedMotion: 'reduce' });
  await page.goto(`/?fixture=${encodeURIComponent(name)}`);
  await expect(page.locator('[data-visual-fixture]')).toHaveAttribute('data-visual-fixture', name);
  await expect(page.locator('[data-tego-sheet]')).toHaveAttribute('data-mode', 'uncontrolled');
  await expect(page.locator('.tego-sheet__canvas')).toBeVisible();
  await expect.poll(() => page.evaluate(() => window.__tegoVisualReady)).toBe(true);
  const fontState = await page.evaluate(async () => {
    await document.fonts.ready;
    return {
      arial: document.fonts.check('400 13px Arial'),
      noto: document.fonts.check('400 13px "Noto Sans Visual"'),
      sourceSans: document.fonts.check('400 12px "Source Sans Pro"'),
      status: document.fonts.status,
    };
  });
  expect(fontState).toEqual({ arial: true, noto: true, sourceSans: true, status: 'loaded' });
}

function expectCoordinate(actual: number, expected: number, label: string): void {
  expect(
    Math.abs(actual - expected),
    `${label}: expected ${expected}, received ${actual}`,
  ).toBeLessThanOrEqual(1);
}

function expectRect(actual: Rect, expected: Rect, label: string): void {
  expectCoordinate(actual.x, expected.x, `${label}.x`);
  expectCoordinate(actual.y, expected.y, `${label}.y`);
  expectCoordinate(actual.width, expected.width, `${label}.width`);
  expectCoordinate(actual.height, expected.height, `${label}.height`);
}

async function recordedRect(
  page: Page,
  record: 'fills' | 'strokeRects',
  color: string,
  expected: Rect,
): Promise<Rect> {
  const candidate = await page.evaluate(
    ({ expectedRect, recordName, targetColor }) => {
      const records =
        recordName === 'fills'
          ? window.__tegoVisual.fills.map((item) => ({ ...item, color: item.fill }))
          : window.__tegoVisual.strokeRects.map((item) => ({ ...item, color: item.stroke }));
      return (
        records
          .filter((item) => item.color === targetColor)
          .map((item) => ({
            distance:
              Math.abs(item.x - expectedRect.x) +
              Math.abs(item.y - expectedRect.y) +
              Math.abs(item.width - expectedRect.width) +
              Math.abs(item.height - expectedRect.height),
            height: item.height,
            width: item.width,
            x: item.x,
            y: item.y,
          }))
          .sort((first, second) => first.distance - second.distance)[0] ?? null
      );
    },
    { expectedRect: expected, recordName: record, targetColor: color },
  );
  expect(candidate, `${record} did not contain ${color}`).not.toBeNull();
  if (candidate === null) throw new Error(`Missing ${color} geometry record`);
  return candidate;
}

async function cellCenter(page: Page, row: number, column: number) {
  const canvas = page.locator('.tego-sheet__canvas');
  const box = await canvas.boundingBox();
  if (box === null) throw new Error('Canvas has no box');
  return { x: box.x + 60 + column * 100 + 50, y: box.y + 25 + row * 25 + 12.5 };
}

async function preparePrintPreview(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet';
    const api = (await import(moduleId)) as typeof import('../../src');
    const parsed = api.parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'visual-print-document',
      workbook: {
        sheets: [
          {
            id: 'visual-print-sheet',
            name: 'Printable',
            cells: [
              {
                row: 0,
                column: 0,
                cell: { input: { type: 'string', value: 'Print report' }, styleId: 'heading' },
              },
              { row: 1, column: 0, cell: { input: { type: 'string', value: 'Visible' } } },
              { row: 1, column: 1, cell: { input: { type: 'number', value: 42 } } },
              {
                row: 1,
                column: 2,
                cell: { input: { type: 'string', value: 'private' }, printable: false },
              },
            ],
            merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 2 } }],
            rows: [{ index: 0, height: 34 }],
            columns: [
              { index: 0, width: 150 },
              { index: 1, width: 120 },
              { index: 2, width: 120 },
            ],
          },
        ],
        styles: [
          {
            id: 'heading',
            value: {
              backgroundColor: '#e8f1ff',
              fontFamily: 'Noto Sans Visual',
              bold: true,
            },
          },
        ],
        validations: [],
        settings: { dateSystem: 'excel-1900', localeHint: 'en-US' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('visual print fixture failed to parse');
    const template = {
      id: 'visual-template',
      name: 'Visual',
      bindings: [],
      printProfiles: [
        {
          id: 'visual-profile',
          name: 'A5 landscape',
          targets: [{ type: 'sheet', sheetId: 'visual-print-sheet' }],
          page: {
            paper: { type: 'A5' },
            orientation: 'landscape',
            margins: { top: 40, right: 40, bottom: 40, left: 40 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: true,
          showHeadings: false,
        },
      ],
    } as unknown as import('../../src').SpreadsheetTemplate;
    const compilation = api.compileSpreadsheetTemplate(parsed.document, template);
    if (!compilation.template) throw new Error('visual template failed to compile');
    const rendered = await api.renderSpreadsheetTemplate(
      {
        template: compilation.template,
        currentDocumentHash: compilation.template.sourceDocumentHash,
        data: {},
        profileId: 'visual-profile',
        missingValue: 'error',
      },
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        clock: new Date('2026-01-01T00:00:00.000Z'),
        fontMetrics: api.createFontMetrics({
          fonts: {
            Arial: { averageAdvance: 7, lineHeight: 13 },
            'Noto Sans Visual': { averageAdvance: 7, lineHeight: 13 },
          },
          fallbackFont: 'Arial',
          fallback: { averageAdvance: 7, lineHeight: 13 },
        }),
      },
    );
    if (!rendered.document) throw new Error('visual template failed to render');
    const pages = api.serializeGeneratedDocumentSvgPages(rendered.document);
    document.querySelector('[data-visual-print-preview]')?.remove();
    const preview = document.createElement('section');
    preview.setAttribute('data-visual-print-preview', '');
    pages.forEach((page, index) => {
      const host = document.createElement('div');
      host.setAttribute('data-visual-print-page', String(index + 1));
      host.style.width = `${page.width}px`;
      host.style.height = `${page.height}px`;
      host.style.lineHeight = '0';
      host.innerHTML = page.svg;
      preview.append(host);
      if (index === 0) {
        const crop = document.createElement('div');
        crop.setAttribute('data-visual-print-crop', 'printable-cells');
        crop.style.width = '410px';
        crop.style.height = '90px';
        crop.style.lineHeight = '0';
        crop.style.overflow = 'hidden';
        crop.innerHTML = page.svg;
        const svg = crop.querySelector('svg');
        svg?.setAttribute('viewBox', '40 40 410 90');
        svg?.setAttribute('width', '410');
        svg?.setAttribute('height', '90');
        preview.append(crop);
      }
    });
    const texts = [...preview.querySelectorAll('text')].map((node) => node.textContent ?? '');
    const fills = [...preview.querySelectorAll<SVGRectElement>('rect[fill]')].map((node) => ({
      fill: node.getAttribute('fill') ?? '',
      height: Number(node.getAttribute('height')),
      width: Number(node.getAttribute('width')),
      x: Number(node.getAttribute('x')),
      y: Number(node.getAttribute('y')),
    }));
    window.__tegoVisual.printSnapshot = {
      css: '@page { size: A5 landscape; }',
      fills,
      pages: pages.map(({ width, height }) => ({ width, height })),
      strokes: preview.querySelectorAll('[stroke]').length,
      texts,
    };
    document.body.append(preview);
  });
  await expect(page.locator('[data-visual-print-page="1"]')).toBeVisible();
}

async function prepareScreenshot(page: Page, fixture: string, touch: boolean): Promise<void> {
  const canvas = page.locator('.tego-sheet__canvas');
  switch (fixture) {
    case 'merged-cells': {
      const point = await cellCenter(page, 0, 0);
      await page.mouse.click(point.x, point.y);
      return;
    }
    case 'frozen-panes': {
      const point = await cellCenter(page, 2, 2);
      await page.mouse.click(point.x, point.y);
      return;
    }
    case 'resized-hidden-structure': {
      await canvas.click({ position: { x: 260, y: 84 } });
      return;
    }
    case 'editing-overlays-menus': {
      await canvas.dblclick({ position: { x: 110, y: 137.5 } });
      const editor = page.getByRole('textbox', { name: 'Cell editor' });
      await expect(editor).toBeVisible();
      await expect(editor).toHaveValue('');
      await editor.fill('Editing');
      await expect(editor).toHaveValue('Editing');
      await editor.evaluate((element) => {
        if (!(element instanceof HTMLTextAreaElement))
          throw new Error('Cell editor is not a textarea');
        element.setSelectionRange(0, 0);
      });
      await page.evaluate(() => window.__tegoVisual.installCaretMask());
      return;
    }
    case 'validation-filter-ui': {
      await page.getByRole('button', { name: 'Filter', exact: true }).click();
      await expect(page.getByRole('dialog', { name: 'Filter' })).toBeVisible();
      return;
    }
    case 'print-preview': {
      await preparePrintPreview(page);
      return;
    }
    case 'touch-interaction': {
      const point = await cellCenter(page, 1, 0);
      if (touch) await page.touchscreen.tap(point.x, point.y);
      else await page.mouse.click(point.x, point.y);
      return;
    }
    default:
      return;
  }
}

test(`${geometryParityToken} geometry gate is green before any screenshot comparison`, async ({
  page,
}) => {
  await openFixture(page, 'default-workbook');
  const canvas = page.locator('.tego-sheet__canvas');
  const canvasBox = await canvas.boundingBox();
  if (canvasBox === null) throw new Error('Canvas has no box');
  const canvasSize = await canvas.evaluate((element) => ({
    height: element.clientHeight,
    width: element.clientWidth,
  }));

  const cell = await recordedRect(page, 'fills', '#ffffff', {
    x: 61,
    y: 26,
    width: 98,
    height: 23,
  });
  expectRect(cell, { x: 61, y: 26, width: 98, height: 23 }, 'cell A1 content');

  const columnHeader = await recordedRect(page, 'fills', '#f4f5f8', {
    x: 0,
    y: 0,
    width: canvasSize.width,
    height: 25,
  });
  expectRect(columnHeader, { x: 0, y: 0, width: canvasSize.width, height: 25 }, 'column header');
  const rowHeader = await recordedRect(page, 'fills', '#f4f5f8', {
    x: 0,
    y: 0,
    width: 60,
    height: canvasSize.height,
  });
  expectRect(rowHeader, { x: 0, y: 0, width: 60, height: canvasSize.height }, 'row header');

  const selection = await recordedRect(page, 'fills', 'rgba(75, 137, 255, 0.1)', {
    x: 60,
    y: 25,
    width: 100,
    height: 25,
  });
  expectRect(selection, { x: 60, y: 25, width: 100, height: 25 }, 'selection');

  await canvas.dblclick({ position: { x: 110, y: 37.5 } });
  const editor = page.locator('.tego-sheet__editor');
  await expect(editor).toBeVisible();
  const editorBox = await editor.boundingBox();
  if (editorBox === null) throw new Error('Editor has no box');
  expectRect(
    {
      height: editorBox.height,
      width: editorBox.width,
      x: editorBox.x - canvasBox.x,
      y: editorBox.y - canvasBox.y,
    },
    {
      x: 60,
      y: 25,
      width: canvasSize.width < 500 ? 158 : 170,
      height: 42.5,
    },
    'editor',
  );
  await page.keyboard.press('Escape');

  const contextPoint = { x: canvasBox.x + 110, y: canvasBox.y + 37.5 };
  await page.mouse.click(contextPoint.x, contextPoint.y, { button: 'right' });
  const overlay = page.getByRole('menu', { name: 'Cell actions' });
  await expect(overlay).toBeVisible();
  const overlayBox = await overlay.boundingBox();
  if (overlayBox === null) throw new Error('Context overlay has no box');
  expectRect(
    {
      height: overlayBox.height,
      width: overlayBox.width,
      x: overlayBox.x - canvasBox.x,
      y: overlayBox.y - canvasBox.y,
    },
    { x: 110, y: 37.5, width: 186, height: 490 },
    'context overlay',
  );

  await openFixture(page, 'frozen-panes');
  const frozenCanvas = page.locator('.tego-sheet__canvas');
  const frozenCanvasSize = await frozenCanvas.evaluate((element) => ({
    height: element.clientHeight,
    width: element.clientWidth,
  }));
  const frozenLines = await page.evaluate(() =>
    window.__tegoVisual.lines.filter((line) => line.stroke === 'rgba(75, 137, 255, 0.6)').slice(-2),
  );
  expect(frozenLines).toHaveLength(2);
  const vertical = frozenLines.find((line) => Math.abs(line.points[0]!.x - line.points[1]!.x) <= 1);
  const horizontal = frozenLines.find(
    (line) => Math.abs(line.points[0]!.y - line.points[1]!.y) <= 1,
  );
  expect(vertical).toBeDefined();
  expect(horizontal).toBeDefined();
  if (vertical === undefined || horizontal === undefined)
    throw new Error('Frozen pane lines are missing');
  expectCoordinate(vertical.points[0]!.x, 160, 'frozen column x');
  expectCoordinate(vertical.points[0]!.y, 25, 'frozen column top');
  expectCoordinate(vertical.points[1]!.y, frozenCanvasSize.height, 'frozen column bottom');
  expectCoordinate(horizontal.points[0]!.x, 60, 'frozen row left');
  expectCoordinate(horizontal.points[1]!.x, frozenCanvasSize.width, 'frozen row right');
  expectCoordinate(horizontal.points[0]!.y, 50, 'frozen row y');
});

for (const fixture of visualFixtures) {
  test(`${visualParityByFixture[fixture.name]} visual fixture: ${fixture.name}`, async ({
    page,
  }, testInfo) => {
    await openFixture(page, fixture.name);
    await prepareScreenshot(page, fixture.name, testInfo.project.name.startsWith('touch-'));
    const screenshot =
      fixture.name === 'print-preview'
        ? page.locator('[data-visual-print-page="1"]')
        : page.locator('[data-tego-sheet]');
    await expect(screenshot).toHaveScreenshot(`${fixture.name}.png`, {
      mask: [...namedMasks(page, fixture.masks ?? [])],
      ...(fixture.name === 'editing-overlays-menus'
        ? {
            maskColor: '#ffffff',
            ...(testInfo.project.name.startsWith('desktop-')
              ? { maxDiffPixels: 50 }
              : { maxDiffPixelRatio: 0 }),
          }
        : {}),
    });
  });
}

test(`${visualParityByFixture['editing-overlays-menus']} context menu from a public right-click interaction`, async ({
  page,
}) => {
  await openFixture(page, 'editing-overlays-menus');
  const point = await cellCenter(page, 4, 0);
  await page.mouse.click(point.x, point.y, { button: 'right' });
  await expect(page.getByRole('menu', { name: 'Cell actions' })).toBeVisible();
  await expect(page.locator('[data-tego-sheet]')).toHaveScreenshot('context-menu.png');
});

test(`${printableCellsParityToken} printable cells are preserved while private cells are omitted`, async ({
  page,
}) => {
  await openFixture(page, 'print-preview');
  await preparePrintPreview(page);
  const printSnapshot = await page.evaluate(() => window.__tegoVisual.printSnapshot);
  expect(printSnapshot).not.toBeNull();
  expect(printSnapshot?.css).toContain('@page { size: A5 landscape; }');
  expect(printSnapshot?.pages).toHaveLength(1);
  expect(printSnapshot?.texts).toContain('Print report');
  expect(printSnapshot?.texts).toContain('Visible');
  expect(printSnapshot?.texts).not.toContain('private');
  expect(printSnapshot?.fills).toContainEqual(
    expect.objectContaining({
      fill: '#e8f1ff',
      height: 32,
      width: 388,
    }),
  );
  expect(printSnapshot?.strokes).toBeGreaterThan(0);
  await expect(page.locator('[data-visual-print-crop="printable-cells"]')).toHaveScreenshot(
    'printable-cells-visual.png',
  );
});
