import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import { parseSpreadsheetDocument } from '../../../src/document';
import {
  bindAdvancedFormula,
  createFormulaNameRegistry,
  planFormulaSpill,
} from '../../../src/formula/advanced';
import { createFormulaEngine } from '../../../src/formula';

describe('FRM-01 advanced formula foundation', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;

  it('binds workbook names and structured references to stable identifiers', () => {
    const names = createFormulaNameRegistry();
    names.register({
      id: 'name-tax',
      name: 'TaxRate',
      scope: 'workbook',
      refersTo: {
        sheetId,
        start: { row: 0, column: 1 },
        end: { row: 0, column: 1 },
      },
    });
    const result = bindAdvancedFormula('=TaxRate+Sales[Amount]', {
      currentSheetId: sheetId,
      names,
      tables: [
        {
          id: 'table-sales',
          name: 'Sales',
          columns: [{ id: 'column-amount', name: 'Amount' }],
        },
      ],
    });
    expect(result.diagnostics).toEqual([]);
    expect(result.references).toEqual([
      { kind: 'name', id: 'name-tax' },
      { kind: 'table-column', tableId: 'table-sales', columnId: 'column-amount' },
    ]);
  });

  it('shadows workbook names with sheet-local names and rejects only same-scope duplicates', () => {
    const names = createFormulaNameRegistry();
    names.register({
      id: 'workbook-rate',
      name: 'Rate',
      scope: 'workbook',
      refersTo: {
        sheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
    });
    names.register({
      id: 'local-rate',
      name: 'RATE',
      scope: { sheetId },
      refersTo: {
        sheetId,
        start: { row: 1, column: 0 },
        end: { row: 1, column: 0 },
      },
    });

    expect(names.resolve('rate', sheetId)?.id).toBe('local-rate');
    expect(names.resolve('rate', 'sheet-2')?.id).toBe('workbook-rate');
    expect(() =>
      names.register({
        id: 'duplicate-local-rate',
        name: 'Rate',
        scope: { sheetId },
        refersTo: {
          sheetId,
          start: { row: 2, column: 0 },
          end: { row: 2, column: 0 },
        },
      }),
    ).toThrowError(expect.objectContaining({ code: 'FORMULA_NAME_CONFLICT' }));
  });

  it('captures named ranges deeply and never binds tokens inside string literals', () => {
    const names = createFormulaNameRegistry();
    const refersTo = {
      sheetId,
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 },
    };
    names.register({
      id: 'name-tax',
      name: 'TaxRate',
      scope: 'workbook',
      refersTo,
    });
    refersTo.start.row = 9;

    expect(names.resolve('TaxRate', sheetId)?.refersTo.start.row).toBe(0);
    expect(Object.isFrozen(names.resolve('TaxRate', sheetId)?.refersTo.start)).toBe(true);
    expect(
      bindAdvancedFormula('="TaxRate"&"Sales[Amount]"', {
        currentSheetId: sheetId,
        names,
        tables: [
          {
            id: 'table-sales',
            name: 'Sales',
            columns: [{ id: 'column-amount', name: 'Amount' }],
          },
        ],
      }),
    ).toEqual({ references: [], diagnostics: [] });
  });

  it('plans dynamic spill atomically and reports blockers', () => {
    expect(
      planFormulaSpill({
        anchor: { sheetId, row: 1, column: 1 },
        value: {
          type: 'array',
          rows: [
            [
              { type: 'number', value: 1 },
              { type: 'number', value: 2 },
            ],
            [
              { type: 'number', value: 3 },
              { type: 'number', value: 4 },
            ],
          ],
        },
        occupied: new Set(['sheet-1!C3']),
        limits: { maxCells: 10 },
      }),
    ).toMatchObject({
      status: 'blocked',
      value: { type: 'error', value: '#SPILL!' },
      blocker: { sheetId: 'sheet-1', row: 2, column: 2 },
    });
    expect(
      planFormulaSpill({
        anchor: { sheetId, row: 1, column: 1 },
        value: {
          type: 'array',
          rows: [[{ type: 'number', value: 1 }], [{ type: 'number', value: 2 }]],
        },
        occupied: new Set(),
        limits: { maxCells: 10 },
      }),
    ).toMatchObject({ status: 'ready', cells: { size: 2 } });
  });

  it('binds names into calculation dependencies and projects dynamic spill values', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'number', value: 5 } } },
              { row: 1, column: 0, cell: { input: { type: 'number', value: 7 } } },
              { row: 0, column: 1, cell: { input: { type: 'formula', source: '=TaxRate*2' } } },
              { row: 0, column: 2, cell: { input: { type: 'formula', source: '=A1:A2' } } },
              { row: 0, column: 3, cell: { input: { type: 'formula', source: '=A1:A2' } } },
              { row: 1, column: 3, cell: { input: { type: 'number', value: 99 } } },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const names = createFormulaNameRegistry();
    names.register({
      id: 'tax-rate',
      name: 'TaxRate',
      scope: 'workbook',
      refersTo: {
        sheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
    });
    const engine = createFormulaEngine({ names });
    const program = engine.compile(parsed.document);
    const result = engine.recalculate(program, [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'number', value: 10 });
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 5 });
    expect(result.values.get('sheet-1!C2')).toEqual({ type: 'number', value: 7 });
    expect(result.values.get('sheet-1!D1')).toEqual({ type: 'error', value: '#SPILL!' });
    expect(program.graph.dependencies.get('sheet-1!B1')).toEqual(new Set(['sheet-1!A1']));

    const unblocked = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 1, column: 3, input: { type: 'blank' } }],
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        clock: { now: () => 0 },
        tick: 1,
        functionRegistryVersion: 'builtin-1',
      },
    );
    expect(unblocked.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 5 });
    expect(unblocked.values.get('sheet-1!D2')).toEqual({ type: 'number', value: 7 });
  });

  it('preserves name scope in dependency cycles and evaluates injected structured references', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'formula', source: '=Loop' } } },
              { row: 0, column: 1, cell: { input: { type: 'number', value: 4 } } },
              { row: 1, column: 1, cell: { input: { type: 'number', value: 6 } } },
              {
                row: 0,
                column: 2,
                cell: { input: { type: 'formula', source: '=SUM(Sales[Amount])' } },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const names = createFormulaNameRegistry();
    names.register({
      id: 'loop',
      name: 'Loop',
      scope: { sheetId },
      refersTo: {
        sheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
    });
    const engine = createFormulaEngine({
      names,
      tables: {
        resolve({ tableName, columnName, currentSheetId }) {
          expect({ tableName, columnName, currentSheetId }).toEqual({
            tableName: 'Sales',
            columnName: 'Amount',
            currentSheetId: 'sheet-1',
          });
          return {
            status: 'resolved',
            tableId: 'table-sales',
            columnId: 'column-amount',
            range: {
              sheetId,
              start: { row: 0, column: 1 },
              end: { row: 1, column: 1 },
            },
          };
        },
      },
    });
    const program = engine.compile(parsed.document);
    const result = engine.recalculate(program, [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.cycles).toEqual([['sheet-1!A1']]);
    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#REF!' });
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 10 });
    expect(program.graph.dependencies.get('sheet-1!C1')).toEqual(
      new Set(['sheet-1!B1', 'sheet-1!B2']),
    );
    expect(program.bindings.get('sheet-1!A1')).toEqual([{ kind: 'name', id: 'loop' }]);
    expect(program.bindings.get('sheet-1!C1')).toEqual([
      {
        kind: 'table-column',
        tableId: 'table-sales',
        columnId: 'column-amount',
      },
    ]);

    const captured = program.bindings.get('sheet-1!C1') as unknown[];
    expect(() => captured.push({ kind: 'poison' })).toThrow(TypeError);
    expect(program.bindings.get('sheet-1!C1')).toEqual([
      {
        kind: 'table-column',
        tableId: 'table-sales',
        columnId: 'column-amount',
      },
    ]);
  });

  it('maps spill-child dependencies to their anchor and cleans projections after anchor edits', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'number', value: 2 } } },
              { row: 1, column: 0, cell: { input: { type: 'number', value: 3 } } },
              { row: 0, column: 2, cell: { input: { type: 'formula', source: '=A1:A2' } } },
              { row: 0, column: 3, cell: { input: { type: 'formula', source: '=C2*2' } } },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const engine = createFormulaEngine();
    const program = engine.compile(parsed.document);
    const environment = {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900' as const,
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    };

    const initial = engine.recalculate(program, [], environment);
    expect(initial.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 6 });
    expect(program.spillAnchors).toEqual(new Map([['sheet-1!C2', 'sheet-1!C1']]));
    expect(program.graph.dependencies.get('sheet-1!D1')).toEqual(new Set(['sheet-1!C1']));
    const unchanged = engine.recalculate(program, [], environment);
    expect(unchanged.evaluatedAddresses).toEqual([]);
    expect(unchanged.values.get('sheet-1!C2')).toEqual({ type: 'number', value: 3 });

    const updated = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 1, column: 0, input: { type: 'number', value: 5 } }],
      { ...environment, tick: 1 },
    );
    expect(updated.evaluatedAddresses).toEqual(['sheet-1!C1', 'sheet-1!D1']);
    expect(updated.values.get('sheet-1!C2')).toEqual({ type: 'number', value: 5 });
    expect(updated.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 10 });

    const removed = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 0, column: 2, input: { type: 'blank' } }],
      { ...environment, tick: 2 },
    );
    expect(removed.values.has('sheet-1!C2')).toBe(false);
    expect(program.spillAnchors).toEqual(new Map());
    expect(removed.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 0 });
  });

  it('blocks deterministic child-only overlap between independent spills', () => {
    const cells = [
      ...Array.from({ length: 2 }, (_, row) =>
        Array.from({ length: 3 }, (_, column) => ({
          row,
          column,
          cell: { input: { type: 'number' as const, value: row * 3 + column + 1 } },
        })),
      ).flat(),
      ...Array.from({ length: 3 }, (_, offset) =>
        Array.from({ length: 2 }, (_, column) => ({
          row: 6 + offset,
          column,
          cell: { input: { type: 'number' as const, value: 10 + offset * 2 + column } },
        })),
      ).flat(),
      { row: 3, column: 2, cell: { input: { type: 'formula' as const, source: '=A1:C2' } } },
      { row: 2, column: 3, cell: { input: { type: 'formula' as const, source: '=A7:B9' } } },
    ];
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [{ id: 'sheet-1', name: 'Sheet 1', cells, merges: [] }],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const engine = createFormulaEngine();
    const program = engine.compile(parsed.document);
    const result = engine.recalculate(program, [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.values.get('sheet-1!C4')).toEqual({ type: 'number', value: 1 });
    expect(result.values.get('sheet-1!D3')).toEqual({ type: 'error', value: '#SPILL!' });
    expect(program.spillAnchors.get('sheet-1!D4')).toBe('sheet-1!C4');
  });

  it('resolves a spill child before an alphabetically later anchor on the initial run', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'formula', source: '=C2*2' } } },
              { row: 0, column: 1, cell: { input: { type: 'number', value: 2 } } },
              { row: 1, column: 1, cell: { input: { type: 'number', value: 3 } } },
              { row: 0, column: 2, cell: { input: { type: 'formula', source: '=B1:B2' } } },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(parsed.document), [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'number', value: 6 });
  });

  it('invalidates child readers when a real input overwrites a spill child', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'number', value: 2 } } },
              { row: 1, column: 0, cell: { input: { type: 'number', value: 3 } } },
              { row: 0, column: 2, cell: { input: { type: 'formula', source: '=A1:A2' } } },
              { row: 0, column: 3, cell: { input: { type: 'formula', source: '=C2*2' } } },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const engine = createFormulaEngine();
    const program = engine.compile(parsed.document);
    const environment = {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900' as const,
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    };
    engine.recalculate(program, [], environment);

    const overwritten = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 1, column: 2, input: { type: 'number', value: 9 } }],
      { ...environment, tick: 1 },
    );

    expect(overwritten.values.get('sheet-1!C1')).toEqual({ type: 'error', value: '#SPILL!' });
    expect(overwritten.values.get('sheet-1!D1')).toEqual({ type: 'number', value: 18 });
    expect(overwritten.evaluatedAddresses).toEqual(['sheet-1!C1', 'sheet-1!D1']);
  });

  it('caches an unchanged blocked spill until its blocker changes', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'number', value: 2 } } },
              { row: 1, column: 0, cell: { input: { type: 'number', value: 3 } } },
              { row: 0, column: 2, cell: { input: { type: 'formula', source: '=A1:A2' } } },
              { row: 1, column: 2, cell: { input: { type: 'number', value: 9 } } },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const engine = createFormulaEngine();
    const program = engine.compile(parsed.document);
    const environment = {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900' as const,
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    };
    engine.recalculate(program, [], environment);

    expect(engine.recalculate(program, [], environment).evaluatedAddresses).toEqual([]);
    const unblocked = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 1, column: 2, input: { type: 'blank' } }],
      { ...environment, tick: 1 },
    );
    expect(unblocked.values.get('sheet-1!C2')).toEqual({ type: 'number', value: 3 });
  });

  it('does not rewrite named or structured-looking text literals during evaluation', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'formula-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'number', value: 5 } } },
              { row: 0, column: 1, cell: { input: { type: 'formula', source: '="TaxRate"' } } },
              {
                row: 0,
                column: 2,
                cell: { input: { type: 'formula', source: '="Sales[Amount]"' } },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('formula fixture must parse');
    const names = createFormulaNameRegistry();
    names.register({
      id: 'tax',
      name: 'TaxRate',
      scope: 'workbook',
      refersTo: {
        sheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
    });
    const engine = createFormulaEngine({
      names,
      tables: {
        resolve: () => ({
          status: 'resolved',
          tableId: 'sales',
          columnId: 'amount',
          range: {
            sheetId,
            start: { row: 0, column: 0 },
            end: { row: 0, column: 0 },
          },
        }),
      },
    });
    const result = engine.recalculate(engine.compile(parsed.document), [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });

    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'string', value: 'TaxRate' });
    expect(result.values.get('sheet-1!C1')).toEqual({
      type: 'string',
      value: 'Sales[Amount]',
    });
  });
});
