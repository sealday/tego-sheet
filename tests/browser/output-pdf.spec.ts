import { expect, test } from '@playwright/test';

test('@parity:output.export-download PDF output runs in the browser main thread and a module Worker', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet/output/pdf';
    const { PdfAdapter } = (await import(moduleId)) as typeof import('../../src/output/pdf');
    const document = {
      workbook: {
        sheets: [],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      print: {
        pages: [
          {
            id: 'browser-pdf',
            index: 0,
            targetId: 'browser',
            width: 96,
            height: 192,
            rowStart: 0,
            rowEnd: 0,
            columnStart: 0,
            columnEnd: 0,
          },
        ],
        displayList: {
          diagnostics: [],
          pages: [
            {
              index: 0,
              width: 96,
              height: 192,
              commands: [
                {
                  kind: 'text',
                  text: 'Worker-safe PDF',
                  x: 8,
                  y: 20,
                  maxWidth: 80,
                  fontFamily: 'Helvetica',
                  fontSize: 12,
                  color: '#000000',
                  horizontalAlign: 'left',
                },
              ],
            },
          ],
        },
      },
      resources: { byHash: {}, byReference: {}, totalBytes: 0 },
      objects: [],
      diagnostics: [],
      metadata: {
        templateId: 'browser',
        profileId: 'browser',
        sourceDocumentHash: 'sha256:browser',
        locale: 'en-US',
        timeZone: 'UTC',
        generatedAt: '2026-07-23T00:00:00.000Z',
      },
    } as unknown as import('../../src').GeneratedDocument;
    const mainBlob = await new PdfAdapter().render(document, { pages: 'all', tagged: false });
    const mainHeader = new TextDecoder().decode((await mainBlob.arrayBuffer()).slice(0, 5));

    const moduleUrl = new URL(moduleId, window.location.href).href;
    const workerSource = `
      self.onmessage = async ({ data }) => {
        try {
          const { PdfAdapter } = await import(${JSON.stringify(moduleUrl)});
          const blob = await new PdfAdapter().render(data, { pages: 'all', tagged: false });
          const bytes = await blob.arrayBuffer();
          self.postMessage({ header: new TextDecoder().decode(bytes.slice(0, 5)), size: blob.size });
        } catch (error) {
          self.postMessage({ error: error instanceof Error ? error.message : String(error) });
        }
      };
    `;
    const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
    try {
      const workerResult = await new Promise<{ header?: string; size?: number; error?: string }>(
        (resolve, reject) => {
          const worker = new Worker(workerUrl, { type: 'module' });
          worker.addEventListener('message', (event) => {
            worker.terminate();
            resolve(event.data as { header?: string; size?: number; error?: string });
          });
          worker.addEventListener('error', reject);
          worker.postMessage(document);
        },
      );
      return { mainHeader, mainSize: mainBlob.size, workerResult };
    } finally {
      URL.revokeObjectURL(workerUrl);
    }
  });

  expect(result.mainHeader).toBe('%PDF-');
  expect(result.mainSize).toBeGreaterThan(500);
  expect(result.workerResult.error).toBeUndefined();
  expect(result.workerResult.header).toBe('%PDF-');
  expect(result.workerResult.size).toBeGreaterThan(500);
});
