import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  type DocumentSheetId,
  type SpreadsheetDocument,
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
          { id: 'region', name: 'Region', dataType: 'text' },
          { id: 'amount', name: 'Amount', dataType: 'number' },
        ],
        headerRows: 1,
        totalsRow: true,
        style: 'CorporateTableStyle',
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

    const exported = await createXlsxWriter().writeResult(parsed.document);
    expect(exported.diagnostics.map(({ details }) => details)).toEqual([
      { feature: 'xlsx:table-column-data-type' },
      { feature: 'xlsx:table-auto-expand' },
    ]);
    await expect(createXlsxWriter().write(parsed.document)).rejects.toMatchObject({
      code: 'DOCUMENT_INVALID',
    });
    const blob = exported.blob;
    const entries = unzipSync(new Uint8Array(await blob.arrayBuffer()));
    expect(strFromU8(entries['xl/worksheets/sheet1.xml']!)).toContain('<tableParts count="1">');
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain(
      'displayName="Sales" ref="A1:B3"',
    );
    expect(strFromU8(entries['xl/tables/table1.xml']!)).toContain('name="CorporateTableStyle"');
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
      style: 'CorporateTableStyle',
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
    await expect(createXlsxWriter().write(parsed.document)).rejects.toMatchObject({
      code: 'DOCUMENT_INVALID',
    });
  });

  it('fails closed when worksheet or table filter state points outside its range', async () => {
    const input = structuredClone(
      createSpreadsheetDocument({ id: 'xlsx-filter-bounds', sheetId: 'sheet-1' }),
    ) as unknown as SpreadsheetDocumentInput;
    const sheet = input.workbook.sheets[0]!;
    sheet.rowCount = 2;
    sheet.columnCount = 2;
    sheet.filter = {
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
      filters: [],
      sort: { column: 2, direction: 'asc' },
    };
    const parsedSheetFilter = parseSpreadsheetDocument(input);
    expect(parsedSheetFilter.ok).toBe(true);
    if (!parsedSheetFilter.ok) throw new Error('sheet filter fixture must parse');
    await expect(createXlsxWriter().writeResult(parsedSheetFilter.document)).rejects.toMatchObject({
      code: 'XLSX_UNSUPPORTED_FEATURE',
    });

    delete sheet.filter;
    sheet.tables = [
      {
        id: 'table-bounds',
        name: 'Bounds',
        range: {
          sheetId: 'sheet-1',
          start: { row: 0, column: 0 },
          end: { row: 1, column: 1 },
        },
        columns: [
          { id: 'first', name: 'First' },
          { id: 'second', name: 'Second' },
        ],
        filter: {
          filters: [{ column: 1, operator: 'in', values: ['inside'] }],
          sort: { column: 1, direction: 'desc' },
        },
      },
    ];
    const parsedTableFilter = parseSpreadsheetDocument(input);
    expect(parsedTableFilter.ok).toBe(true);
    if (!parsedTableFilter.ok) throw new Error('table filter fixture must parse');
    const parsedTable = parsedTableFilter.document.workbook.sheets[0]!.tables[0]!;
    const unsafeTableFilter = {
      ...parsedTableFilter.document,
      workbook: {
        ...parsedTableFilter.document.workbook,
        sheets: [
          {
            ...parsedTableFilter.document.workbook.sheets[0]!,
            tables: [
              {
                ...parsedTable,
                filter: {
                  filters: [{ column: 2, operator: 'in', values: ['outside'] }],
                  sort: { column: 2, direction: 'desc' },
                },
              },
            ],
          },
        ],
      },
    } as SpreadsheetDocument;
    await expect(createXlsxWriter().writeResult(unsafeTableFilter)).rejects.toMatchObject({
      code: 'XLSX_UNSUPPORTED_FEATURE',
    });
  });

  it('rejects imported table filter and sort state outside the table range', async () => {
    const base = createSpreadsheetDocument({ id: 'xlsx-table-bounds', sheetId: 'sheet-1' });
    const input = structuredClone(base) as unknown as SpreadsheetDocumentInput;
    const sheet = input.workbook.sheets[0]!;
    sheet.rowCount = 2;
    sheet.columnCount = 2;
    sheet.tables = [
      {
        id: 'table-bounds',
        name: 'Bounds',
        range: {
          sheetId: 'sheet-1',
          start: { row: 0, column: 0 },
          end: { row: 1, column: 1 },
        },
        columns: [
          { id: 'first', name: 'First' },
          { id: 'second', name: 'Second' },
        ],
      },
    ];
    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error('table bounds fixture must parse');
    const exported = await createXlsxWriter().writeResult(parsed.document);
    const entries = unzipSync(new Uint8Array(await exported.blob.arrayBuffer()));
    const tableXml = strFromU8(entries['xl/tables/table1.xml']!);
    const invalidPackage = (replacement: string) =>
      zipSync({
        ...entries,
        'xl/tables/table1.xml': strToU8(
          tableXml.replace('<autoFilter ref="A1:B2"></autoFilter>', replacement),
        ),
      });

    await expect(
      createXlsxReader().read(
        invalidPackage(
          '<autoFilter ref="A1:B2"><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>',
        ),
      ),
    ).rejects.toMatchObject({ code: 'MALFORMED_WORKBOOK' });
    await expect(
      createXlsxReader().read(
        invalidPackage(
          '<autoFilter ref="A1:B2"><sortState ref="A1:B2"><sortCondition ref="C2:C2"/></sortState></autoFilter>',
        ),
      ),
    ).rejects.toMatchObject({ code: 'MALFORMED_WORKBOOK' });
  });
});
