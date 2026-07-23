import { expect, test } from '@playwright/test';

test('XLSX output runs deterministically in the browser and a module Worker', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet/output/xlsx';
    const { XlsxAdapter } = (await import(moduleId)) as typeof import('../../src/output/xlsx');
    const document = {
      workbook: {
        sheets: [
          {
            id: 'sheet',
            name: 'Worker',
            cells: [{ row: 0, column: 0, cell: { input: { type: 'string', value: '=literal' } } }],
            merges: [],
            rows: [],
            columns: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      calculatedCells: [],
      print: {
        pages: [],
        displayList: { diagnostics: [], pages: [] },
        profile: {
          id: 'worker',
          name: 'Worker',
          targets: [{ type: 'sheet', sheetId: 'sheet' }],
          page: {
            paper: { type: 'A4' },
            orientation: 'portrait',
            margins: { top: 0, right: 0, bottom: 0, left: 0 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: true,
          showHeadings: true,
        },
      },
      resources: { byHash: {}, byReference: {}, totalBytes: 0 },
      objects: [],
      diagnostics: [],
      metadata: {
        templateId: 'worker',
        profileId: 'worker',
        sourceDocumentHash: 'sha256:worker',
        locale: 'en-US',
        timeZone: 'UTC',
        generatedAt: '2026-07-23T00:00:00.000Z',
      },
    } as unknown as import('../../src').GeneratedDocument;
    const options = {
      formulaMode: 'values-only' as const,
      compatibility: 'excel' as const,
    };
    const adapter = new XlsxAdapter();
    const first = new Uint8Array(await (await adapter.render(document, options)).arrayBuffer());
    const second = new Uint8Array(await (await adapter.render(document, options)).arrayBuffer());
    const equal =
      first.length === second.length && first.every((value, index) => value === second[index]);

    const moduleUrl = new URL(moduleId, window.location.href).href;
    const workerSource = `
      self.onmessage = async ({ data }) => {
        try {
          const { XlsxAdapter } = await import(${JSON.stringify(moduleUrl)});
          const blob = await new XlsxAdapter().render(data, {
            formulaMode: 'values-only',
            compatibility: 'excel',
          });
          const bytes = new Uint8Array(await blob.arrayBuffer());
          self.postMessage({ header: Array.from(bytes.slice(0, 4)), size: blob.size });
        } catch (error) {
          self.postMessage({ error: error instanceof Error ? error.message : String(error) });
        }
      };
    `;
    const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
    try {
      const workerResult = await new Promise<{
        header?: readonly number[];
        size?: number;
        error?: string;
      }>((resolve, reject) => {
        const worker = new Worker(workerUrl, { type: 'module' });
        worker.addEventListener('message', (event) => {
          worker.terminate();
          resolve(event.data as { header?: readonly number[]; size?: number; error?: string });
        });
        worker.addEventListener('error', reject);
        worker.postMessage(document);
      });
      return { equal, header: Array.from(first.slice(0, 4)), workerResult };
    } finally {
      URL.revokeObjectURL(workerUrl);
    }
  });

  expect(result.equal).toBe(true);
  expect(result.header).toEqual([80, 75, 3, 4]);
  expect(result.workerResult.error).toBeUndefined();
  expect(result.workerResult.header).toEqual([80, 75, 3, 4]);
  expect(result.workerResult.size).toBeGreaterThan(500);
});
