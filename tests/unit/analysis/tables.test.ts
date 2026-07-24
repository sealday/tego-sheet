import { describe, expect, it, vi } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type BindingId,
  type DocumentSheetId,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createFormulaEngine } from '../../../src/formula';
import {
  createStructuredTableResolver,
  executeStructuredTableView,
  planStructuredTableAutoExpand,
} from '../../../src/analysis/tables';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';
import { createSheetGridModel } from '../../../src/engine';

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

function selection(
  sheet: ReturnType<typeof sheetId>,
  row: number,
  column: number,
  endRow = row,
  endColumn = column,
) {
  return {
    sheet,
    active: { row, column },
    range: {
      start: { row, column },
      end: { row: endRow, column: endColumn },
    },
  } as const;
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

  it('persists table presentation, behavior, filter, and typed-column metadata', () => {
    const document = parseOk(
      fixture([
        {
          ...fixture().workbook.sheets[0]!.tables![0]!,
          columns: [
            { id: 'column-region', name: 'Region', dataType: 'text' },
            { id: 'column-amount', name: 'Amount', dataType: 'number' },
          ],
          headerRows: 1,
          totalsRow: true,
          style: 'medium-2',
          autoExpand: true,
          filter: {
            filters: [{ column: 1, operator: 'in', values: ['4', '6'] }],
            sort: { column: 1, direction: 'desc' },
          },
        },
      ]),
    );
    const roundTrip = parseOk(JSON.parse(serializeSpreadsheetDocument(document)));

    expect(roundTrip.workbook.sheets[0]?.tables[0]).toMatchObject({
      headerRows: 1,
      totalsRow: true,
      style: 'medium-2',
      autoExpand: true,
      filter: {
        filters: [{ column: 1, operator: 'in', values: ['4', '6'] }],
        sort: { column: 1, direction: 'desc' },
      },
    });
    expect(
      createStructuredTableResolver(roundTrip).resolve({
        tableName: 'Sales',
        columnName: 'Amount',
        currentSheetId: 'sheet-1',
      }),
    ).toMatchObject({ range: { start: { row: 1 }, end: { row: 1 } } });
  });

  it('executes stable table-local filtering and sorting against one immutable revision', () => {
    const table = {
      ...parseOk(fixture()).workbook.sheets[0]!.tables[0]!,
      filter: {
        filters: [{ column: 0, operator: 'in' as const, values: ['East', 'West'] }],
        sort: { column: 1, direction: 'desc' as const },
      },
    };
    const values = new Map<string, unknown>([
      ['1:0', 'East'],
      ['1:1', 4],
      ['2:0', 'West'],
      ['2:1', 6],
    ]);

    expect(
      executeStructuredTableView(table, {
        revision: 'table-r1',
        read: (row, column) => values.get(`${row}:${column}`),
      }),
    ).toEqual({ sourceRevision: 'table-r1', rowIndices: [2, 1] });
    expect(() =>
      executeStructuredTableView(
        table,
        { revision: 'table-r1', read: () => undefined },
        { maximumRows: 1 },
      ),
    ).toThrowError(/row limit exceeded/iu);
  });

  it('uses the structured table projection in the production screen grid', () => {
    const input = fixture([
      {
        ...fixture().workbook.sheets[0]!.tables![0]!,
        filter: {
          filters: [{ column: 0, operator: 'in', values: ['West'] }],
          sort: { column: 1, direction: 'desc' },
        },
      },
    ]);
    const controller = new SpreadsheetDocumentController(parseOk(input));
    const model = createSheetGridModel(controller.getSnapshot().projection[0]!);

    expect(Array.from({ length: 3 }, (_, visual) => model.logicalRowAtVisualIndex(visual))).toEqual(
      [0, 2, 1],
    );
    expect(model.rowHeight(1)).toBe(0);
    expect(model.rowHeight(2)).toBeGreaterThan(0);
  });

  it.each([
    {
      name: 'filter',
      filter: {
        filters: [{ column: 2, operator: 'in' as const, values: ['East'] }],
      },
    },
    {
      name: 'sort',
      filter: {
        filters: [],
        sort: { column: 2, direction: 'asc' as const },
      },
    },
  ])(
    'rejects a table $name column outside the table range during parsing and execution',
    ({ filter }) => {
      const input = fixture([
        {
          ...fixture().workbook.sheets[0]!.tables![0]!,
          filter,
        },
      ]);
      const parsed = parseSpreadsheetDocument(input);
      expect(parsed.ok).toBe(false);
      if (parsed.ok) throw new Error('out-of-range table filter metadata must fail');
      expect(parsed.diagnostics).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            code: 'INVALID_RANGE',
            message: expect.stringMatching(/column.*table range/iu),
          }),
        ]),
      );

      const table = {
        ...parseOk(fixture()).workbook.sheets[0]!.tables[0]!,
        filter,
      };
      const read = vi.fn();
      expect(() => executeStructuredTableView(table, { revision: 'table-r1', read })).toThrowError(
        /column.*table range/iu,
      );
      expect(read).not.toHaveBeenCalled();

      const controller = new SpreadsheetDocumentController(parseOk(fixture()));
      expect(() =>
        controller.dispatch({ type: 'set-table', sheet: sheetId('sheet-1'), table }, 'ref'),
      ).toThrowError(/column.*table range/iu);
      expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.filter).toBeUndefined();
    },
  );

  it('plans direct bounded auto expansion and rejects occupied rows atomically', () => {
    const table = {
      ...parseOk(fixture()).workbook.sheets[0]!.tables[0]!,
      autoExpand: true,
    };
    const append = { sheetId: table.range.sheetId, row: 3, column: 1 };

    expect(planStructuredTableAutoExpand(table, append)).toMatchObject({
      status: 'expanded',
      table: { range: { end: { row: 3, column: 1 } } },
    });
    expect(
      planStructuredTableAutoExpand(table, append, [
        {
          sheetId: table.range.sheetId,
          start: { row: 3, column: 0 },
          end: { row: 3, column: 1 },
        },
      ]),
    ).toEqual({ status: 'rejected', code: 'TABLE_RANGE_OVERLAP' });
    expect(planStructuredTableAutoExpand(table, append, [], 4)).toEqual({
      status: 'rejected',
      code: 'TABLE_CELL_LIMIT_EXCEEDED',
    });
  });

  it.each([
    {
      name: 'set-cell-text',
      dispatch(controller: SpreadsheetDocumentController) {
        return controller.dispatch(
          {
            type: 'set-cell-text',
            address: { sheet: sheetId('sheet-1'), row: 3, column: 1 },
            text: '8',
          },
          'ref',
        );
      },
    },
    {
      name: 'set-cell-input',
      dispatch(controller: SpreadsheetDocumentController) {
        return controller.dispatch(
          {
            type: 'set-cell-input',
            address: { sheet: sheetId('sheet-1'), row: 3, column: 1 },
            input: { type: 'number', value: 8 },
          },
          'ref',
        );
      },
    },
    {
      name: 'paste-external',
      dispatch(controller: SpreadsheetDocumentController) {
        return controller.dispatch(
          {
            type: 'paste-external',
            target: selection(sheetId('sheet-1'), 3, 0),
            values: [['North', '8']],
          },
          'ref',
        );
      },
    },
    {
      name: 'paste-internal',
      dispatch(controller: SpreadsheetDocumentController) {
        return controller.dispatch(
          {
            type: 'paste-internal',
            source: selection(sheetId('sheet-1'), 2, 0, 2, 1),
            target: selection(sheetId('sheet-1'), 3, 0, 3, 1),
            mode: 'value',
            cut: false,
          },
          'ref',
        );
      },
    },
    {
      name: 'autofill',
      dispatch(controller: SpreadsheetDocumentController) {
        return controller.dispatch(
          {
            type: 'autofill',
            source: selection(sheetId('sheet-1'), 2, 0, 2, 1),
            target: selection(sheetId('sheet-1'), 3, 0, 3, 1),
            mode: 'value',
          },
          'ref',
        );
      },
    },
  ])('auto-expands $name in the same undoable transaction', ({ dispatch }) => {
    const controller = new SpreadsheetDocumentController(
      parseOk(
        fixture([
          {
            ...fixture().workbook.sheets[0]!.tables![0]!,
            autoExpand: true,
          },
        ]),
      ),
    );

    expect(dispatch(controller).status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.range.end).toEqual({
      row: 3,
      column: 1,
    });
    expect(controller.undo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.range.end).toEqual({
      row: 2,
      column: 1,
    });
    expect(controller.redo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.tables[0]?.range.end.row).toBe(3);
  });

  it('rejects template repeat and table auto-expand ownership of the same boundary', () => {
    const input = fixture([
      {
        ...fixture().workbook.sheets[0]!.tables![0]!,
        autoExpand: true,
      },
    ]);
    input.templates.push({
      id: 'template-table',
      name: 'Table report',
      bindings: [
        {
          id: 'repeat-table' as BindingId,
          type: 'repeat-rows',
          range: {
            sheetId: 'sheet-1' as DocumentSheetId,
            start: { row: 1, column: 0 },
            end: { row: 2, column: 1 },
          },
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
      ],
      printProfiles: [],
    });

    const result = parseSpreadsheetDocument(input);
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('boundary conflict must be rejected');
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'TABLE_TEMPLATE_BOUNDARY_CONFLICT' }),
      ]),
    );
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

  it('rewrites structured-reference display names by stable IDs when a table is renamed', () => {
    const input = fixture();
    input.workbook.sheets.push({
      id: 'sheet-qualified',
      name: 'Sales[Amount]',
      rowCount: 1,
      columnCount: 1,
      cells: [{ row: 0, column: 0, cell: { input: { type: 'number', value: 1 } } }],
      merges: [],
      tables: [],
    });
    const controller = new SpreadsheetDocumentController(parseOk(input));
    const table = controller.getDocument().workbook.sheets[0]!.tables[0]!;
    controller.dispatch(
      {
        type: 'set-cell-input',
        address: { sheet: sheetId('sheet-1'), row: 1, column: 3 },
        input: {
          type: 'formula',
          source: `="Sales[Amount]"&'Sales[Amount]'!A1+SUM(Sales[Amount])`,
        },
      },
      'ref',
    );

    expect(
      controller.dispatch(
        {
          type: 'set-table',
          sheet: sheetId('sheet-1'),
          table: {
            ...table,
            name: 'Revenue',
            columns: [table.columns[0]!, { ...table.columns[1]!, name: 'NetAmount' }],
          },
        },
        'ref',
      ).status,
    ).toBe('committed');

    expect(controller.getDocument().workbook.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          row: 0,
          column: 3,
          cell: { input: { type: 'formula', source: '=SUM(Revenue[NetAmount])' } },
        }),
        expect.objectContaining({
          row: 1,
          column: 3,
          cell: {
            input: {
              type: 'formula',
              source: `="Sales[Amount]"&'Sales[Amount]'!A1+SUM(Revenue[NetAmount])`,
            },
          },
        }),
      ]),
    );
    const program = createFormulaEngine({
      tables: createStructuredTableResolver(() => controller.getDocument()),
    }).compile(controller.getDocument());
    expect(program.bindings.get('sheet-1!D1')).toEqual([
      { kind: 'table-column', tableId: 'table-sales', columnId: 'column-amount' },
    ]);

    expect(controller.undo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          cell: { input: { type: 'formula', source: '=SUM(Sales[Amount])' } },
        }),
      ]),
    );
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
