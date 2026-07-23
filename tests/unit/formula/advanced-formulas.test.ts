import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import {
  bindAdvancedFormula,
  createFormulaNameRegistry,
  planFormulaSpill,
} from '../../../src/formula/advanced';

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
});
