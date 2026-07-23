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
});
