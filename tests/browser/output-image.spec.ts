import { expect, test } from '@playwright/test';

test('@parity:output.export-download Image output emits inert SVG and a structurally valid PNG at the requested pixels', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleId = '/@id/tego-sheet/output/image';
    const { ImageAdapter } = (await import(moduleId)) as typeof import('../../src/output/image');
    const document = {
      workbook: {
        sheets: [],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      calculatedCells: [],
      print: {
        pages: [
          {
            id: 'browser-image',
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
                  kind: 'fill-rect',
                  rect: { x: 0, y: 0, width: 96, height: 192 },
                  color: '#ff0000',
                },
                {
                  kind: 'link',
                  rect: { x: 1, y: 1, width: 4, height: 4 },
                  href: 'https://example.com/removed',
                  label: 'external link',
                },
              ],
            },
          ],
        },
        profile: {
          id: 'browser',
          name: 'Browser',
          targets: [],
          page: {
            paper: { type: 'custom', width: 96, height: 192 },
            orientation: 'portrait',
            margins: { top: 0, right: 0, bottom: 0, left: 0 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: false,
          showHeadings: false,
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

    const adapter = new ImageAdapter();
    const [svgBlob] = await adapter.render(document, {
      format: 'svg',
      pages: [0],
      background: 'transparent',
    });
    const svg = await svgBlob!.text();
    const [pngBlob] = await adapter.render(document, {
      format: 'png',
      pages: [0],
      dpi: 150,
      background: 'transparent',
    });
    const bytes = new Uint8Array(await pngBlob!.arrayBuffer());
    const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const chunks: string[] = [];
    let offset = 8;
    while (offset + 12 <= bytes.length) {
      const length = data.getUint32(offset);
      chunks.push(String.fromCharCode(...bytes.slice(offset + 4, offset + 8)));
      offset += 12 + length;
    }

    const bitmap = await createImageBitmap(pngBlob!);
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    const context = canvas.getContext('2d')!;
    context.drawImage(bitmap, 0, 0);
    const pixel = Array.from(context.getImageData(75, 150, 1, 1).data);
    bitmap.close();

    return {
      svg,
      signature: Array.from(bytes.slice(0, 8)),
      ihdr: {
        length: data.getUint32(8),
        type: String.fromCharCode(...bytes.slice(12, 16)),
        width: data.getUint32(16),
        height: data.getUint32(20),
        bitDepth: bytes[24],
        colorType: bytes[25],
      },
      chunks,
      decoded: { width: canvas.width, height: canvas.height, pixel },
    };
  });

  expect(result.svg).toContain('viewBox="0 0 96 192"');
  expect(result.svg).not.toMatch(/<script|<foreignObject|\s(?:src|href|xlink:href)=["']https?:/iu);
  expect(result.signature).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect(result.ihdr).toEqual({
    length: 13,
    type: 'IHDR',
    width: 150,
    height: 300,
    bitDepth: 8,
    colorType: 6,
  });
  expect(result.chunks[0]).toBe('IHDR');
  expect(result.chunks).toContain('IDAT');
  expect(result.chunks.at(-1)).toBe('IEND');
  expect(result.decoded).toEqual({
    width: 150,
    height: 300,
    pixel: [255, 0, 0, 255],
  });
});

test('@parity:output.export-download PNG output rasterizes inside a module Worker', async ({
  browserName,
  page,
}) => {
  if (browserName !== 'chromium') return;
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const moduleUrl = new URL('/@id/tego-sheet/output/image', window.location.href).href;
    const workerSource = `
      self.onmessage = async ({ data }) => {
        try {
          const { ImageAdapter } = await import(${JSON.stringify(moduleUrl)});
          const [blob] = await new ImageAdapter().render(data, {
            format: 'png',
            pages: [0],
            dpi: 150,
            background: 'transparent',
          });
          const bytes = new Uint8Array(await blob.arrayBuffer());
          const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
          const bitmap = await createImageBitmap(blob);
          self.postMessage({
            signature: Array.from(bytes.slice(0, 8)),
            width: view.getUint32(16),
            height: view.getUint32(20),
            decodedWidth: bitmap.width,
            decodedHeight: bitmap.height,
          });
          bitmap.close();
        } catch (error) {
          self.postMessage({ error: error instanceof Error ? error.message : String(error) });
        }
      };
    `;
    const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
    const document = {
      workbook: {
        sheets: [],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      calculatedCells: [],
      print: {
        pages: [
          {
            id: 'worker-image',
            index: 0,
            targetId: 'worker',
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
                  kind: 'fill-rect',
                  rect: { x: 0, y: 0, width: 96, height: 192 },
                  color: '#ff0000',
                },
              ],
            },
          ],
        },
        profile: {
          id: 'worker',
          name: 'Worker',
          targets: [],
          page: {
            paper: { type: 'custom', width: 96, height: 192 },
            orientation: 'portrait',
            margins: { top: 0, right: 0, bottom: 0, left: 0 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: false,
          showHeadings: false,
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
    };
    try {
      return await new Promise<{
        signature?: readonly number[];
        width?: number;
        height?: number;
        decodedWidth?: number;
        decodedHeight?: number;
        error?: string;
      }>((resolve, reject) => {
        const worker = new Worker(workerUrl, { type: 'module' });
        worker.addEventListener('message', (event) => {
          worker.terminate();
          resolve(event.data);
        });
        worker.addEventListener('error', reject);
        worker.postMessage(document);
      });
    } finally {
      URL.revokeObjectURL(workerUrl);
    }
  });

  expect(result.error).toBeUndefined();
  expect(result.signature).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect(result).toMatchObject({
    width: 150,
    height: 300,
    decodedWidth: 150,
    decodedHeight: 300,
  });
});
