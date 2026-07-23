import { describe, expect, it } from 'vitest';
import { createFormulaEngine } from '../../../src/formula/evaluator';
import type { CalculationEnvironment } from '../../../src/formula/evaluator';
import { formulaDocument } from './helpers';

const environment: CalculationEnvironment = {
  locale: 'en-US',
  timeZone: 'UTC',
  dateSystem: 'excel-1900',
  clock: { now: () => 1_700_000_000_000 },
  tick: 0,
  functionRegistryVersion: 'builtin-1',
};

describe('formula dependency graph', () => {
  it('recalculates only transitive dirty dependents in stable order', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 1 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1+1' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=B1+1' } },
          { row: 0, column: 25, input: { type: 'formula', source: '=1+1' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    engine.recalculate(program, [{ sheetId: 'sheet-1', row: 0, column: 0 }], environment);
    const result = engine.recalculate(
      program,
      [{ sheetId: 'sheet-1', row: 0, column: 0 }],
      environment,
    );
    expect(result.evaluatedAddresses).toEqual(['sheet-1!B1', 'sheet-1!C1']);
    expect(result.evaluatedAddresses).not.toContain('sheet-1!Z1');
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 3 });
  });

  it('returns a stable minimal cycle and propagates standard errors', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=B1' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=1/0' } },
          { row: 0, column: 3, input: { type: 'formula', source: '=C1+1' } },
          { row: 0, column: 4, input: { type: 'formula', source: '=NOPE(1)' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(result.cycles).toEqual([['sheet-1!A1', 'sheet-1!B1']]);
    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#REF!' });
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(result.values.get('sheet-1!D1')).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(result.values.get('sheet-1!E1')).toEqual({ type: 'error', value: '#NAME?' });
  });

  it('uses the explicit clock and invalidates volatile functions only on tick changes', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=NOW()' } }],
      },
    ]);
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    const first = engine.recalculate(program, [], environment);
    const unchanged = engine.recalculate(program, [], environment);
    const next = engine.recalculate(program, [], { ...environment, tick: 1 });
    expect(first.values.get('sheet-1!A1')).toEqual(unchanged.values.get('sheet-1!A1'));
    expect(unchanged.evaluatedAddresses).toEqual([]);
    expect(next.evaluatedAddresses).toEqual(['sheet-1!A1']);
  });
});
