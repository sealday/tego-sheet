import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  type DocumentSheetId,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createXlsxReader, createXlsxWriter } from '../../../src/interchange';

describe('XLSX native structured tables', () => {
  it('writes native table parts and reads their supported semantics back', async () => {
    const base = createSpreadsheetDocument({ id: 'xlsx-table', sheetId: 'sheet-1' });
    const input = structuredClone(base) as unknown as SpreadsheetDocumentInput;
    const sheet = input.workbook.sheets[0]!;
    sheet.rowCount = 3;
    sheet.columnCount = 2;
    sheet.cells = [
      { row: 0, column: 0, cell: { input: { type: 'string', value: 'Region' } } },
      { row: 0, column: 1, cell: { input: { type: 'string', value: 'Amount' } } },
      { row: 1, column: 0, cell: { input: { type: 'string', value: 'East' } } },
      { row: 1, column: 1, cell: { input: { type: 'number', value: 4 } } },
      { row: 2, column: 0, cell: { input: { type: 'string', value: 'Total' } } },
      { row: 2, column: 1, cell: { input: { type: 'number', value: 4 } } },
    ];
    sheet.tables = [
      {
        id: 'table-sales',
        name: 'Sales',
        range: {
          sheetId: 'sheet-1',
          start: { row: 0, column: 0 },
          end: { row: 2, column: 1 },
        },
        columns: [
          { id: 'region', name: 'Region' },
          { id: 'amount', name: 'Amount' },
        ],
        headerRows: 1,
        totalsRow: true,
        style: 'TableStyleMedium4',
        autoExpand: true,
        filter: {
          filters: [{ column: 0, operator: 'in', values: ['East'] }],
          sort: { column: 1, direction: 'desc' },
        },
      },
    ];
    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error('table fixture must parse');

    const blob = await createXlsxWriter().write(parsed.document);
    const entries = unzipSync(new Uint8Array(await blob.arrayBuffer()));
    expect(strFromU8(entries['xl/worksheets/sheet1.xml']!)).toContain('<tableParts count="1">');
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain(
      'displayName="Sales" ref="A1:B3"',
    );
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain('name="TableStyleMedium4"');
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain(
      '<filterColumn colId="0"><filters><filter val="East"/>',
    );
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain('descending="1"');

    const imported = await createXlsxReader().read(blob);
    expect(imported.security.unsupportedFeatures).not.toContain('xlsx:tables');
    expect(imported.document.workbook.sheets[0]?.tables[0]).toMatchObject({
      name: 'Sales',
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
      columns: [{ name: 'Region' }, { name: 'Amount' }],
      totalsRow: true,
      style: 'TableStyleMedium4',
      autoExpand: true,
      filter: {
        filters: [{ column: 0, operator: 'in', values: ['East'] }],
        sort: { column: 1, direction: 'desc' },
      },
    });
  });

  it('reports explicit chart and sparkline degradation instead of silently dropping them', async () => {
    const sheetId = 'sheet-1' as DocumentSheetId;
    const input = structuredClone(
      createSpreadsheetDocument({ id: 'xlsx-visuals', sheetId: 'sheet-1' }),
    ) as unknown as SpreadsheetDocumentInput;
    const sheet = input.workbook.sheets[0]!;
    sheet.charts = [
      {
        id: 'chart-1',
        type: 'line',
        series: [
          {
            id: 'series-1',
            values: {
              sheetId,
              start: { row: 0, column: 0 },
              end: { row: 0, column: 1 },
            },
          },
        ],
      },
    ];
    sheet.sparklines = [
      {
        id: 'spark-1',
        type: 'line',
        source: {
          sheetId,
          start: { row: 0, column: 0 },
          end: { row: 0, column: 1 },
        },
        target: { sheetId, row: 0, column: 2 },
      },
    ];
    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error('visual fixture must parse');

    const result = await createXlsxWriter().writeResult!(parsed.document);
    expect(result.diagnostics.map(({ details }) => details)).toEqual([
      { feature: 'xlsx:charts' },
      { feature: 'xlsx:sparklines' },
    ]);
  });
});
