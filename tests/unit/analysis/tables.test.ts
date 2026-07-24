import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createFormulaEngine } from '../../../src/formula';
import { createStructuredTableResolver } from '../../../src/analysis/tables';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';

function fixture(
  tables: SpreadsheetDocumentInput['workbook']['sheets'][number]['tables'] = [
    {
      id: 'table-sales',
      name: 'Sales',
      range: {
        sheetId: 'sheet-1',
        start: { row: 0, column: 0 },
        end: { row: 2, column: 1 },
      },
      columns: [
        { id: 'column-region', name: 'Region' },
        { id: 'column-amount', name: 'Amount' },
      ],
    },
  ],
): SpreadsheetDocumentInput {
  return {
    schemaVersion: 2,
    id: 'tables-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Data',
          rowCount: 10,
          columnCount: 5,
          cells: [
            { row: 0, column: 0, cell: { input: { type: 'string', value: 'Region' } } },
            { row: 0, column: 1, cell: { input: { type: 'string', value: 'Amount' } } },
            { row: 1, column: 0, cell: { input: { type: 'string', value: 'East' } } },
            { row: 1, column: 1, cell: { input: { type: 'number', value: 4 } } },
            { row: 2, column: 0, cell: { input: { type: 'string', value: 'West' } } },
            { row: 2, column: 1, cell: { input: { type: 'number', value: 6 } } },
            {
              row: 0,
              column: 3,
              cell: { input: { type: 'formula', source: '=SUM(Sales[Amount])' } },
            },
          ],
          merges: [],
          tables,
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  };
}

function parseOk(input: SpreadsheetDocumentInput): SpreadsheetDocument {
  const result = parseSpreadsheetDocument(input);
  expect(result.ok, JSON.stringify(result.diagnostics)).toBe(true);
  if (!result.ok) throw new Error('fixture must parse');
  return result.document;
}

describe('TBL-01 persistent structured tables', () => {
  it('round-trips stable table and column IDs in canonical ID order', () => {
    const input = fixture([
      {
        id: 'table-z',
        name: 'Zulu',
        range: {
          sheetId: 'sheet-1',
          start: { row: 0, column: 3 },
          end: { row: 2, column: 4 },
        },
        columns: [
          { id: 'column-z1', name: 'First' },
          { id: 'column-z2', name: 'Second' },
        ],
      },
      {
        id: 'table-a',
        name: 'Alpha',
        range: {
          sheetId: 'sheet-1',
          start: { row: 4, column: 0 },
          end: { row: 6, column: 1 },
        },
        columns: [
          { id: 'column-a1', name: 'First' },
          { id: 'column-a2', name: 'Second' },
        ],
      },
    ]);

    const first = parseOk(input);
    const second = parseOk(JSON.parse(serializeSpreadsheetDocument(first)));

    expect(second.workbook.sheets[0]?.tables.map(({ id }) => id)).toEqual(['table-a', 'table-z']);
    expect(second.workbook.sheets[0]?.tables[0]?.columns).toEqual([
      { id: 'column-a1', name: 'First' },
      { id: 'column-a2', name: 'Second' },
    ]);
  });

  it.each([
    {
      name: 'case-insensitive duplicate workbook table names',
      tables: [
        ...fixture().workbook.sheets[0]!.tables!,
        {
          id: 'table-duplicate',
          name: 'sales',
          range: {
            sheetId: 'sheet-1',
            start: { row: 4, column: 0 },
            end: { row: 6, column: 1 },
          },
          columns: [
            { id: 'duplicate-a', name: 'A' },
            { id: 'duplicate-b', name: 'B' },
          ],
        },
      ],
    },
    {
      name: 'overlapping table ranges',
      tables: [
        ...fixture().workbook.sheets[0]!.tables!,
        {
          id: 'table-overlap',
          name: 'Other',
          range: {
            sheetId: 'sheet-1',
            start: { row: 2, column: 1 },
            end: { row: 4, column: 2 },
          },
          columns: [
            { id: 'overlap-a', name: 'A' },
            { id: 'overlap-b', name: 'B' },
          ],
        },
      ],
    },
    {
      name: 'duplicate column names',
      tables: [
        {
          ...fixture().workbook.sheets[0]!.tables![0]!,
          columns: [
            { id: 'column-region', name: 'Amount' },
            { id: 'column-amount', name: 'amount' },
          ],
        },
      ],
    },
    {
      name: 'column counts that do not match range width',
      tables: [
        {
          ...fixture().workbook.sheets[0]!.tables![0]!,
          columns: [{ id: 'column-region', name: 'Region' }],
        },
      ],
    },
  ])('rejects $name atomically', ({ tables }) => {
    const result = parseSpreadsheetDocument(fixture(tables));
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('invalid table fixture must fail');
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: expect.stringMatching(/TABLE|DUPLICATE|RANGE/),
        }),
      ]),
    );
  });

  it('enforces table and column decode budgets before materializing the document', () => {
    const input = fixture();
    expect(parseSpreadsheetDocument(input, { limits: { maxTables: 0 } })).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'TABLE_LIMIT_EXCEEDED' })],
    });
    expect(parseSpreadsheetDocument(input, { limits: { maxTableColumns: 1 } })).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'TABLE_LIMIT_EXCEEDED' })],
    });
  });

  it('rejects stable column ID reuse across different tables', () => {
    const result = parseSpreadsheetDocument(
      fixture([
        ...fixture().workbook.sheets[0]!.tables!,
        {
          id: 'table-other',
          name: 'Other',
          range: {
            sheetId: 'sheet-1',
            start: { row: 4, column: 0 },
            end: { row: 6, column: 1 },
          },
          columns: [
            { id: 'column-region', name: 'OtherRegion' },
            { id: 'column-other', name: 'OtherAmount' },
          ],
        },
      ]),
    );

    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('reused stable column ID must fail');
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'DUPLICATE_ID',
          message: 'Duplicate stable ID column-region',
        }),
      ]),
    );
  });

  it('resolves persisted columns to data rows and evaluates structured references', () => {
    const document = parseOk(fixture());
    const resolver = createStructuredTableResolver(document);

    expect(
      resolver.resolve({
        tableName: 'sales',
        columnName: 'AMOUNT',
        currentSheetId: 'sheet-1',
      }),
    ).toEqual({
      status: 'resolved',
      tableId: 'table-sales',
      columnId: 'column-amount',
      range: {
        sheetId: 'sheet-1',
        start: { row: 1, column: 1 },
        end: { row: 2, column: 1 },
      },
    });

    const engine = createFormulaEngine({ tables: resolver });
    const program = engine.compile(document);
    const result = engine.recalculate(program, [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 10 });
    expect(program.bindings.get('sheet-1!D1')).toEqual([
      { kind: 'table-column', tableId: 'table-sales', columnId: 'column-amount' },
    ]);
  });

  it('commits table changes atomically and restores them through undo and redo', () => {
    const empty = fixture([]);
    const controller = new SpreadsheetDocumentController(parseOk(empty));
    const table = parseOk(fixture()).workbook.sheets[0]!.tables[0]!;

    expect(
      controller.dispatch({ type: 'set-table', sheet: sheetId('sheet-1'), table }, 'ref').status,
    ).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables).toEqual([table]);

    expect(controller.undo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables).toEqual([]);

    expect(controller.redo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables).toEqual([table]);

    expect(
      controller.dispatch(
        { type: 'remove-table', sheet: sheetId('sheet-1'), tableId: 'table-sales' },
        'ref',
      ).status,
    ).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables).toEqual([]);
  });

  it('tracks row structure changes but rejects column edits that would invent column identities', () => {
    const controller = new SpreadsheetDocumentController(parseOk(fixture()));

    expect(
      controller.dispatch(
        { type: 'insert-row', sheet: sheetId('sheet-1'), index: 2, count: 1 },
        'ref',
      ).status,
    ).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.range).toEqual({
      sheetId: 'sheet-1',
      start: { row: 0, column: 0 },
      end: { row: 3, column: 1 },
    });

    expect(() =>
      controller.dispatch(
        { type: 'insert-column', sheet: sheetId('sheet-1'), index: 1, count: 1 },
        'ref',
      ),
    ).toThrowError(/structured table column identities/u);
    expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.columns).toHaveLength(2);
  });
});
