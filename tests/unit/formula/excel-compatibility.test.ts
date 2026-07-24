import { describe, expect, it } from 'vitest';
import {
  BUILTIN_FORMULA_COMPATIBILITY,
  createFormulaFunctionRegistry,
} from '../../../src/formula/function-registry';
import { createFormulaEngine } from '../../../src/formula/evaluator';
import { formulaDocument } from './helpers';

const environment = {
  locale: 'en-US',
  timeZone: 'UTC',
  dateSystem: 'excel-1900' as const,
  clock: { now: () => Date.UTC(2024, 0, 2, 3, 4, 5) },
  tick: 1,
  functionRegistryVersion: 'builtin-1',
};

describe('Excel compatibility manifest', () => {
  it('declares every built-in function and evaluates the supported matrix', () => {
    expect(BUILTIN_FORMULA_COMPATIBILITY.map(({ name }) => name)).toEqual([
      'ABS',
      'AND',
      'AVERAGE',
      'COUNT',
      'COUNTA',
      'COUNTBLANK',
      'CONCAT',
      'DATE',
      'DAY',
      'IF',
      'LEFT',
      'LEN',
      'LOWER',
      'MAX',
      'MID',
      'MIN',
      'MONTH',
      'NOT',
      'NOW',
      'OR',
      'RIGHT',
      'ROUND',
      'ROUNDDOWN',
      'ROUNDUP',
      'SUM',
      'TRIM',
      'TODAY',
      'UPPER',
      'YEAR',
    ]);
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=SUM(1,2,3)' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=IF(1=1,7,8)' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=TODAY()' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'number', value: 6 });
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'number', value: 7 });
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 45293 });
  });

  it('evaluates the deterministic numeric, text, logical, date, and counting baseline', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=ABS(-3)' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=ROUND(2.345,2)' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=ROUNDUP(2.341,2)' } },
          { row: 0, column: 3, input: { type: 'formula', source: '=ROUNDDOWN(2.349,2)' } },
          { row: 0, column: 4, input: { type: 'formula', source: '=LEN(TRIM("  a   b "))' } },
          { row: 0, column: 5, input: { type: 'formula', source: '=UPPER(LEFT("sales",2))' } },
          { row: 0, column: 6, input: { type: 'formula', source: '=LOWER(RIGHT("QTY",2))' } },
          { row: 0, column: 7, input: { type: 'formula', source: '=MID("amount",2,3)' } },
          { row: 0, column: 8, input: { type: 'formula', source: '=NOT(OR(FALSE,0))' } },
          { row: 0, column: 9, input: { type: 'formula', source: '=DATE(2024,2,29)' } },
          { row: 0, column: 10, input: { type: 'formula', source: '=YEAR(J1)' } },
          { row: 0, column: 11, input: { type: 'formula', source: '=MONTH(J1)' } },
          { row: 0, column: 12, input: { type: 'formula', source: '=DAY(J1)' } },
          { row: 1, column: 0, input: { type: 'number', value: 1 } },
          { row: 2, column: 0, input: { type: 'string', value: 'x' } },
          { row: 3, column: 0, input: { type: 'formula', source: '=COUNT(A2:A3)' } },
          { row: 4, column: 0, input: { type: 'formula', source: '=COUNTA(A2:A3)' } },
          { row: 5, column: 0, input: { type: 'formula', source: '=COUNTBLANK(B2:B3)' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(
      Array.from({ length: 13 }, (_, column) =>
        result.values.get(`sheet-1!${String.fromCharCode(65 + column)}1`),
      ),
    ).toEqual([
      { type: 'number', value: 3 },
      { type: 'number', value: 2.35 },
      { type: 'number', value: 2.35 },
      { type: 'number', value: 2.34 },
      { type: 'number', value: 3 },
      { type: 'string', value: 'SA' },
      { type: 'string', value: 'ty' },
      { type: 'string', value: 'mou' },
      { type: 'boolean', value: true },
      { type: 'number', value: 45351 },
      { type: 'number', value: 2024 },
      { type: 'number', value: 2 },
      { type: 'number', value: 29 },
    ]);
    expect(result.values.get('sheet-1!A4')).toEqual({ type: 'number', value: 1 });
    expect(result.values.get('sheet-1!A5')).toEqual({ type: 'number', value: 2 });
    expect(result.values.get('sheet-1!A6')).toEqual({ type: 'number', value: 2 });
  });

  it('returns stable errors for invalid arguments and enforces a calculation-step budget', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=ROUND("x",2)' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=1+2+3+4' } },
        ],
      },
    ]);
    const engine = createFormulaEngine({ maximumCalculationSteps: 5 });
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#VALUE!' });
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'error', value: '#NUM!' });
    expect(result.diagnostics.get('sheet-1!B1')).toEqual([
      expect.objectContaining({ code: 'FORMULA_EVALUATION_LIMIT_EXCEEDED' }),
    ]);
  });

  it('registers frozen, explicitly typed extension functions without global access', () => {
    const registry = createFormulaFunctionRegistry();
    registry.register({
      name: 'DOUBLE',
      parameters: { minimum: 1, maximum: 1 },
      returns: 'number',
      volatility: 'stable',
      mode: 'sync',
      evaluate: ([value], context) => {
        expect(Object.isFrozen(context)).toBe(true);
        return { type: 'number', value: Number(value?.type === 'number' ? value.value : 0) * 2 };
      },
    });
    expect(registry.version).toBe('builtin-1+1');
    expect(registry.resolve('DOUBLE')?.name).toBe('DOUBLE');
  });

  it('derives volatile date serials from the explicit time zone', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=TODAY()' } }],
      },
    ]);
    const clock = { now: () => Date.UTC(2024, 0, 1, 20) };
    const utcEngine = createFormulaEngine();
    const shanghaiEngine = createFormulaEngine();
    expect(
      utcEngine
        .recalculate(utcEngine.compile(document), [], {
          ...environment,
          clock,
          timeZone: 'UTC',
        })
        .values.get('sheet-1!A1'),
    ).toEqual({ type: 'number', value: 45292 });
    expect(
      shanghaiEngine
        .recalculate(shanghaiEngine.compile(document), [], {
          ...environment,
          clock,
          timeZone: 'Asia/Shanghai',
        })
        .values.get('sheet-1!A1'),
    ).toEqual({ type: 'number', value: 45293 });
  });

  it('applies the Excel DATE rule that treats years 0 through 1899 as 1900-based', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=YEAR(DATE(100,1,1))' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=YEAR(DATE(0,1,1))' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);

    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'number', value: 2000 });
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'number', value: 1900 });
  });
});
