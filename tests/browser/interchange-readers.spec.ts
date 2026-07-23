import { expect, test } from '@playwright/test';

test('@parity:interchange.readers CSV, TSV, XLSX, and ODS readers run in the browser main thread and a module Worker', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet';
    const api = (await import(moduleId)) as typeof import('../../src');
    const fixtures = await Promise.all(
      (
        [
          ['csv', '/fixtures/interchange/minimal.csv'],
          ['tsv', '/fixtures/interchange/minimal.tsv'],
          ['xlsx', '/fixtures/interchange/minimal.xlsx'],
          ['ods', '/fixtures/interchange/minimal.ods'],
        ] as const
      ).map(async ([format, url]) => {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Fixture ${url} returned ${response.status}`);
        return { format, bytes: await response.arrayBuffer() };
      }),
    );

    const reader = (format: (typeof fixtures)[number]['format']) => {
      switch (format) {
        case 'csv':
          return api.createCsvReader();
        case 'tsv':
          return api.createTsvReader();
        case 'xlsx':
          return api.createXlsxReader();
        case 'ods':
          return api.createOdsReader();
      }
    };
    const main = await Promise.all(
      fixtures.map(async ({ format, bytes }) => {
        const imported = await reader(format).read(bytes);
        return {
          format: imported.format,
          first: imported.document.workbook.sheets[0]?.cells[0]?.cell.input,
          security: imported.security,
        };
      }),
    );

    const moduleUrl = new URL('/src/interchange-worker.ts', window.location.href).href;
    const workerSource = `
      self.onmessage = async ({ data }) => {
        try {
          const api = await import(${JSON.stringify(moduleUrl)});
          const reader = (format) => {
            switch (format) {
              case 'csv': return api.createCsvReader();
              case 'tsv': return api.createTsvReader();
              case 'xlsx': return api.createXlsxReader();
              case 'ods': return api.createOdsReader();
            }
          };
          const results = await Promise.all(data.map(async ({ format, bytes }) => {
            const imported = await reader(format).read(bytes);
            return {
              format: imported.format,
              first: imported.document.workbook.sheets[0]?.cells[0]?.cell.input,
              security: imported.security,
            };
          }));
          self.postMessage({ results });
        } catch (error) {
          self.postMessage({ error: error instanceof Error ? error.message : String(error) });
        }
      };
    `;
    const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
    try {
      const worker = await new Promise<{ results?: typeof main; error?: string }>(
        (resolve, reject) => {
          const instance = new Worker(workerUrl, { type: 'module' });
          instance.addEventListener('message', (event) => {
            instance.terminate();
            resolve(event.data as { results?: typeof main; error?: string });
          });
          instance.addEventListener('error', reject);
          instance.postMessage(fixtures);
        },
      );
      return { main, worker };
    } finally {
      URL.revokeObjectURL(workerUrl);
    }
  });

  const expected = [
    { format: 'csv', first: { type: 'string', value: 'name' } },
    { format: 'tsv', first: { type: 'string', value: 'name' } },
    { format: 'xlsx', first: { type: 'string', value: 'Hello from XLSX' } },
    { format: 'ods', first: { type: 'string', value: 'Hello from ODS' } },
  ];
  expect(result.main).toMatchObject(expected);
  expect(result.main.every(({ security }) => security.activeContentExecuted === false)).toBe(true);
  expect(result.main.every(({ security }) => security.externalResourcesFetched === false)).toBe(
    true,
  );
  expect(result.worker.error).toBeUndefined();
  expect(result.worker.results).toEqual(result.main);
});
