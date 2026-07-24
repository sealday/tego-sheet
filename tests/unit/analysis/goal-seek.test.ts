import { describe, expect, it } from 'vitest';
import { solveFormulaGoalSeek } from '../../../src/analysis/goal-seek';
import { createFormulaFunctionRegistry } from '../../../src/formula';
import { formulaDocument } from '../formula/helpers';

const environment = {
  locale: 'en-US',
  timeZone: 'UTC',
  dateSystem: 'excel-1900' as const,
  clock: { now: () => 0 },
  tick: 0,
  functionRegistryVersion: createFormulaFunctionRegistry().version,
};

describe('formula goal seek', () => {
  it('solves in an isolated program and evaluates only transitive dependents', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 1 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1*2' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=B1+1' } },
          { row: 0, column: 3, input: { type: 'formula', source: '=40+2' } },
        ],
      },
    ]);

    const result = solveFormulaGoalSeek({
      document,
      variable: { sheetId: 'sheet-1', row: 0, column: 0 },
      target: { sheetId: 'sheet-1', row: 0, column: 2 },
      targetValue: 10,
      environment,
    });

    expect(result.status).toBe('converged');
    expect(result.value).toBeCloseTo(4.5, 8);
    expect(result.targetValue).toBeCloseTo(10, 8);
    expect(new Set(result.evaluatedAddresses)).toEqual(new Set(['sheet-1!B1', 'sheet-1!C1']));
    expect(document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'number',
      value: 1,
    });
  });

  it('honors iteration budgets and returns its best bounded result', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 0 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1*3' } },
        ],
      },
    ]);

    const result = solveFormulaGoalSeek({
      document,
      variable: { sheetId: 'sheet-1', row: 0, column: 0 },
      target: { sheetId: 'sheet-1', row: 0, column: 1 },
      targetValue: 30,
      initialGuess: 0,
      initialStep: 1,
      maximumIterations: 1,
      environment,
    });

    expect(result.status).toBe('iteration-limit');
    expect(result.iterations).toBe(1);
  });

  it('fails closed for cancellation and non-numeric targets', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 1 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=IF(A1,"text","text")' } },
        ],
      },
    ]);
    const controller = new AbortController();
    controller.abort();

    expect(
      solveFormulaGoalSeek({
        document,
        variable: { sheetId: 'sheet-1', row: 0, column: 0 },
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
        targetValue: 2,
        signal: controller.signal,
        environment,
      }).status,
    ).toBe('cancelled');

    expect(() =>
      solveFormulaGoalSeek({
        document,
        variable: { sheetId: 'sheet-1', row: 0, column: 0 },
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
        targetValue: 2,
        environment,
      }),
    ).toThrow(/numeric/u);
  });

  it('rejects targets that are not dependent on the variable', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 1 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=4+2' } },
        ],
      },
    ]);

    expect(() =>
      solveFormulaGoalSeek({
        document,
        variable: { sheetId: 'sheet-1', row: 0, column: 0 },
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
        targetValue: 10,
        environment,
      }),
    ).toThrow(/does not depend/u);
  });
});
