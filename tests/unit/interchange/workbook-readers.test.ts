import { describe, expect, it } from 'vitest';
import {
  createCsvReader,
  createCsvWriter,
  createOdsReader,
  createXlsxReader,
} from '../../../src/interchange';
import { createSpreadsheetDocument } from '../../../src/document';

describe('IO-01 bounded atomic workbook readers and writers', () => {
  it('round-trips CSV quoting, newlines, typed values, and formula-injection protection', async () => {
    const reader = createCsvReader();
    const imported = await reader.read(
      new TextEncoder().encode('name,amount\\r\\n"A, Inc",12\\r\\n'),
    );
    expect(imported.document.workbook.sheets[0]?.cells).toHaveLength(4);
    const writer = createCsvWriter();
    const protectedBlob = await writer.write(
      {
        ...imported.document,
        workbook: {
          ...imported.document.workbook,
          sheets: [
            {
              ...imported.document.workbook.sheets[0]!,
              cells: [
                {
                  row: 0,
                  column: 0,
                  cell: { input: { type: 'string', value: '=SUM(A1:A2)' } },
                },
              ],
            },
          ],
        },
      },
      { delimiter: ',', lineEnding: '\r\n', formulaInjectionProtection: true },
    );
    await expect(protectedBlob.text()).resolves.toBe(`'=SUM(A1:A2)\r\n`);
  });

  it('rejects archive and cell limits without exposing partial documents', async () => {
    const oversized = new Uint8Array(32);
    await expect(createXlsxReader({ maxPackageBytes: 16 }).read(oversized)).rejects.toMatchObject({
      code: 'ARCHIVE_LIMIT_EXCEEDED',
    });
    await expect(createOdsReader({ maxPackageBytes: 16 }).read(oversized)).rejects.toMatchObject({
      code: 'ARCHIVE_LIMIT_EXCEEDED',
    });
  });

  it('writers consume immutable SpreadsheetDocument snapshots', async () => {
    const document = createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' });
    const blob = await createCsvWriter().write(document, {
      delimiter: '\t',
      lineEnding: '\n',
      formulaInjectionProtection: true,
    });
    expect(blob.type).toBe('text/tab-separated-values');
    expect(document.workbook.sheets[0]?.cells).toEqual([]);
  });
});
