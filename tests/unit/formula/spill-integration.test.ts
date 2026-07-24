import { describe, expect, it } from 'vitest';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';
import {
  createFormulaEngine,
  createFormulaFunctionRegistry,
  createFormulaNameRegistry,
  formulaAddressKey,
} from '../../../src/formula';
import { formulaDocument } from './helpers';
import { createPresentationCache, createPresentationResolver } from '../../../src/presentation';

function spillDocument() {
  return formulaDocument([
    {
      id: 'sheet-1',
      name: 'Sheet1',
      cells: [
        { row: 0, column: 0, input: { type: 'number', value: 2 } },
        { row: 1, column: 0, input: { type: 'number', value: 3 } },
        { row: 0, column: 1, input: { type: 'formula', source: '=A1:A2' } },
      ],
    },
  ]);
}

describe('FRM-01 spill production integration', () => {
  it('projects child values through presentation and controller snapshots', () => {
    const document = spillDocument();
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    engine.recalculate(program, [], {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: { now: () => 0 },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });
    const resolver = createPresentationResolver({
      document,
      formulaProgram: program,
      cache: createPresentationCache({ maximumEntries: 10, maximumBytes: 10_000 }),
      revisions: {
        document: 0,
        calculation: 0,
        condition: 0,
        style: 0,
        environment: 0,
      },
      environment: {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        target: 'screen',
      },
    });
    expect(resolver.resolve({ sheetId: 'sheet-1' as never, row: 1, column: 1 })).toMatchObject({
      value: { type: 'number', value: 3 },
      formattedText: '3',
      accessibility: { readOnly: true },
    });

    const controller = new SpreadsheetDocumentController(document, {
      calculation: { functions: createFormulaFunctionRegistry() },
    });
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 1, column: 1 })).toBe('');
    expect(
      controller.getSnapshot().calculation.values.find(({ address }) => address === 'sheet-1!B2')
        ?.value,
    ).toEqual({ type: 'number', value: 3 });
  });

  it('rejects non-anchor edits and restores spill projection through undo', () => {
    const controller = new SpreadsheetDocumentController(spillDocument(), {
      calculation: { functions: createFormulaFunctionRegistry() },
    });
    const sheet = sheetId('sheet-1');
    const rejected = controller.execute({
      schemaVersion: 1,
      id: 'edit-spill-child',
      command: {
        type: 'set-cell-text',
        address: { sheet, row: 1, column: 1 },
        text: 'blocked',
      },
    });
    expect(rejected).toMatchObject({
      status: 'rejected',
      code: 'SPILL_CELL_READ_ONLY',
    });
    expect(controller.getCellText({ sheet, row: 1, column: 1 })).toBe('');

    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 1 }, text: '9' },
      'ref',
    );
    expect(controller.getCellText({ sheet, row: 1, column: 1 })).toBe('');
    expect(
      controller
        .getDocument()
        .workbook.sheets[0]?.cells.some(({ row, column }) => row === 1 && column === 1),
    ).toBe(false);
    controller.undo();
    expect(controller.getCellText({ sheet, row: 1, column: 1 })).toBe('');
    expect(
      controller
        .getSnapshot()
        .calculation.values.some(
          ({ address, value }) =>
            address === formulaAddressKey({ sheetId: 'sheet-1', row: 1, column: 1 }) &&
            value.type === 'number' &&
            value.value === 3,
        ),
    ).toBe(true);
  });

  it('rejects merges that intersect an active spill child', () => {
    const controller = new SpreadsheetDocumentController(spillDocument(), {
      calculation: { functions: createFormulaFunctionRegistry() },
    });
    const sheet = sheetId('sheet-1');
    const result = controller.execute({
      schemaVersion: 1,
      id: 'merge-spill-child',
      command: {
        type: 'merge',
        selection: {
          sheet,
          range: { start: { row: 0, column: 1 }, end: { row: 1, column: 1 } },
          active: { row: 0, column: 1 },
        },
      },
    });

    expect(result).toMatchObject({ status: 'rejected', code: 'SPILL_CELL_READ_ONLY' });
  });

  it('passes names and structured-table providers through the production controller', () => {
    const names = createFormulaNameRegistry();
    names.register({
      id: 'tax',
      name: 'TaxRate',
      scope: 'workbook',
      refersTo: {
        sheetId: 'sheet-1' as never,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
    });
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 4 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=TaxRate*2' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=SUM(Sales[Amount])' } },
        ],
      },
    ]);
    const controller = new SpreadsheetDocumentController(document, {
      calculation: {
        functions: createFormulaFunctionRegistry(),
        names,
        tables: {
          resolve: () => ({
            status: 'resolved',
            tableId: 'sales',
            columnId: 'amount',
            range: {
              sheetId: 'sheet-1' as never,
              start: { row: 0, column: 0 },
              end: { row: 0, column: 0 },
            },
          }),
        },
      },
    });
    const values = new Map(
      controller.getSnapshot().calculation.values.map(({ address, value }) => [address, value]),
    );

    expect(values.get('sheet-1!B1')).toEqual({ type: 'number', value: 8 });
    expect(values.get('sheet-1!C1')).toEqual({ type: 'number', value: 4 });
  });

  it('clears stale parse errors and diagnostics when a formula becomes text', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=1+' } }],
      },
    ]);
    const controller = new SpreadsheetDocumentController(document, {
      calculation: { functions: createFormulaFunctionRegistry() },
    });
    const sheet = sheetId('sheet-1');
    expect(
      controller.getSnapshot().calculation.values.find(({ address }) => address === 'sheet-1!A1')
        ?.value,
    ).toEqual({ type: 'error', value: '#VALUE!' });

    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'plain' },
      'ref',
    );

    expect(
      controller.getSnapshot().calculation.values.some(({ address }) => address === 'sheet-1!A1'),
    ).toBe(false);
    expect(controller.getCellText({ sheet, row: 0, column: 0 })).toBe('plain');
  });
});
