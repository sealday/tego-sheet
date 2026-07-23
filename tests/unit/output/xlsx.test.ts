import { describe, expect, it } from 'vitest';
import { XlsxAdapter } from '../../../src/output/xlsx';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

async function parts(blob: Blob): Promise<Readonly<Record<string, string>>> {
  const { strFromU8, unzipSync } = await import('fflate');
  const archive = unzipSync(new Uint8Array(await blob.arrayBuffer()));
  return Object.fromEntries(
    Object.entries(archive).map(([path, bytes]) => [path, strFromU8(bytes)]),
  );
}

describe('XlsxAdapter', () => {
  it('writes typed cells, literal equals text, formulas with caches, styles, merges, validation, dimensions, and print parts', async () => {
    const source = {
      ...outputGeneratedDocument(),
      calculatedValues: {
        'sheet-1:1:1': { type: 'number', value: 3 },
      },
      print: {
        ...outputGeneratedDocument().print,
        profile: {
          id: 'invoice-print',
          name: 'Invoice',
          targets: [{ type: 'sheet', sheetId: 'sheet-1' }],
          page: {
            paper: { type: 'A4' },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          repeatRows: {
            sheetId: 'sheet-1',
            start: { row: 0, column: 0 },
            end: { row: 0, column: 1 },
          },
          manualBreaks: [{ sheetId: 'sheet-1', beforeRow: 2 }],
          header: { center: 'Invoice' },
          footer: { right: 'Page &P' },
          showGridlines: false,
          showHeadings: true,
        },
      },
    };
    const blob = await new XlsxAdapter().render(source, {
      formulaMode: 'formula-and-cached-value',
      compatibility: 'excel',
    });
    const xml = await parts(blob);
    const sheet = xml['xl/worksheets/sheet1.xml'] ?? '';

    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(sheet).toContain('<c r="A1"');
    expect(sheet).toContain('<t>=literal</t>');
    expect(sheet).toContain('<f>A2+1</f><v>3</v>');
    expect(sheet).toContain('<mergeCell ref="A1:B1"/>');
    expect(sheet).toContain('<dataValidation');
    expect(sheet).toContain('<col min="2" max="2" width="80" hidden="1"');
    expect(sheet).toContain('<row r="1" ht="24" customHeight="1"');
    expect(sheet).toContain('<pageMargins');
    expect(sheet).toContain('<pageSetup');
    expect(sheet).toContain('<rowBreaks');
    expect(sheet).toContain('<headerFooter>');
    expect(xml['xl/styles.xml']).toContain('<numFmt');
    expect(xml['xl/workbook.xml']).toContain('_xlnm.Print_Titles');
  });

  it('removes formula markup in values-only mode', async () => {
    const xml = await parts(
      await new XlsxAdapter().render(outputGeneratedDocument(), {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    );

    expect(xml['xl/worksheets/sheet1.xml']).not.toContain('<f>');
  });

  it('produces byte-identical ZIP output for equal inputs and options', async () => {
    const adapter = new XlsxAdapter();
    const options = {
      formulaMode: 'values-only' as const,
      compatibility: 'excel' as const,
    };

    const first = new Uint8Array(
      await (await adapter.render(outputGeneratedDocument(), options)).arrayBuffer(),
    );
    const second = new Uint8Array(
      await (await adapter.render(outputGeneratedDocument(), options)).arrayBuffer(),
    );
    expect(second).toEqual(first);
  });

  it('rejects unsupported custom cells with a located diagnostic and no package', async () => {
    const source = outputGeneratedDocument();
    const custom = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: {
                    type: 'custom',
                    cellType: 'vendor/widget',
                    schemaVersion: 1,
                    value: {},
                  },
                },
              },
            ],
          },
        ],
      },
    };

    await expect(
      new XlsxAdapter().render(custom as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({
      code: 'XLSX_UNSUPPORTED_FEATURE',
      diagnostic: { location: { sheetId: 'sheet-1' } },
    });
  });

  it('aborts before publishing a partial ZIP', async () => {
    const controller = new AbortController();
    controller.abort();

    await expect(
      new XlsxAdapter().render(outputGeneratedDocument(), {
        formulaMode: 'values-only',
        compatibility: 'excel',
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
  });
});
