import { expect, test } from '@playwright/test';

test('@parity:output.browser-print prints generated SVG pages in an isolated iframe and cleans after afterprint', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet';
    const api = (await import(moduleId)) as typeof import('../../src');
    const document = api.createSpreadsheetDocument({
      id: 'browser-print-document',
      sheetId: 'browser-print-sheet',
      sheetName: 'Invoice',
    });
    const template = {
      id: 'browser-template',
      name: 'Invoice',
      bindings: [],
      printProfiles: [
        {
          id: 'browser-profile',
          name: 'Browser',
          targets: [{ type: 'sheet', sheetId: 'browser-print-sheet' }],
          page: {
            paper: { type: 'custom', width: 200, height: 300 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: true,
          showHeadings: false,
        },
      ],
    } as unknown as import('../../src').SpreadsheetTemplate;
    const compilation = api.compileSpreadsheetTemplate(document, template);
    if (!compilation.template) throw new Error('compile failed');
    const rendered = await api.renderSpreadsheetTemplate(
      {
        template: compilation.template,
        currentDocumentHash: compilation.template.sourceDocumentHash,
        data: {},
        profileId: 'browser-profile',
        missingValue: 'error',
      },
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        clock: new Date('2026-01-01T00:00:00.000Z'),
        fontMetrics: api.createFontMetrics({
          fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
          fallbackFont: 'Arial',
          fallback: { averageAdvance: 6, lineHeight: 12 },
        }),
      },
    );
    if (!rendered.document) throw new Error('render failed');
    let isolated = false;
    const adapter = new api.IsolatedBrowserPrintAdapter({
      print(target: Window) {
        isolated =
          target.document.querySelector('[data-tego-sheet], [data-tego-toolbar]') === null &&
          target.document.querySelectorAll('svg[data-page-id]').length ===
            rendered.document!.print.pages.length;
        target.dispatchEvent(new Event('afterprint'));
      },
    });
    const printed = await adapter.print(rendered.document);
    adapter.dispose();
    return {
      isolated,
      pageCount: printed.pageCount,
      remainingFrames: window.document.querySelectorAll('[data-tego-browser-print]').length,
    };
  });

  expect(result.isolated).toBe(true);
  expect(result.pageCount).toBeGreaterThan(0);
  expect(result.remainingFrames).toBe(0);
});
