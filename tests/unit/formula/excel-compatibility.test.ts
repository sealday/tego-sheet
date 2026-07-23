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
      'AND',
      'AVERAGE',
      'CONCAT',
      'IF',
      'MAX',
      'MIN',
      'NOW',
      'OR',
      'SUM',
      'TODAY',
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
});
