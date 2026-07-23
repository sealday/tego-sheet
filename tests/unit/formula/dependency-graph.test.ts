import { describe, expect, it } from 'vitest';
import { createFormulaEngine } from '../../../src/formula/evaluator';
import type { CalculationEnvironment } from '../../../src/formula/evaluator';
import { createFormulaFunctionRegistry } from '../../../src/formula/function-registry';
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
  it('does not expose mutable dependency state through program snapshots', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 1 } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1+1' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    engine.recalculate(program, [], environment);
    (program.graph.dependencies as Map<string, ReadonlySet<string>>).clear();
    (program.graph.dependents as Map<string, ReadonlySet<string>>).clear();
    const result = engine.recalculate(
      program,
      [
        {
          sheetId: 'sheet-1',
          row: 0,
          column: 0,
          input: { type: 'number', value: 2 },
        },
      ],
      environment,
    );
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'number', value: 3 });
  });

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

    const edited = engine.recalculate(
      program,
      [
        {
          sheetId: 'sheet-1',
          row: 0,
          column: 0,
          input: { type: 'number', value: 5 },
        },
      ],
      environment,
    );
    expect(edited.evaluatedAddresses).toEqual(['sheet-1!B1', 'sheet-1!C1']);
    expect(edited.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 7 });
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

  it('propagates errors through ranges and lazily evaluates IF branches', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=1/0' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=SUM(A1:A1)' } },
          { row: 0, column: 2, input: { type: 'formula', source: '=IF(TRUE,1,1/0)' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(result.values.get('sheet-1!C1')).toEqual({ type: 'number', value: 1 });
  });

  it('enforces dependency expansion limits before publishing a program', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=SUM(A1:A100001)' } }],
      },
    ]);
    expect(() => createFormulaEngine().compile(document)).toThrow(/dependency limit/u);

    const cumulative = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          {
            row: 0,
            column: 0,
            input: { type: 'formula', source: '=SUM(A1:A60000,A60001:A120000)' },
          },
        ],
      },
    ]);
    expect(() => createFormulaEngine().compile(cumulative)).toThrow(/dependency limit/u);
  });

  it('does not mutate a program when an incremental graph rebuild exceeds its budget', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=1' } }],
      },
    ]);
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    engine.recalculate(program, [], environment);

    expect(() =>
      engine.recalculate(
        program,
        [
          {
            sheetId: 'sheet-1',
            row: 0,
            column: 0,
            input: { type: 'formula', source: '=SUM(A1:A100001)' },
          },
        ],
        environment,
      ),
    ).toThrow(/dependency limit/u);

    expect(program.formulas.get('sheet-1!A1')).toMatchObject({ kind: 'number', value: 1 });
    expect(program.values.get('sheet-1!A1')).toEqual({ type: 'number', value: 1 });
  });

  it('propagates an incrementally introduced parse error and clears stale cycle diagnostics', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=1' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=A1+1' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const program = engine.compile(document);
    engine.recalculate(program, [], environment);
    const invalid = engine.recalculate(
      program,
      [
        {
          sheetId: 'sheet-1',
          row: 0,
          column: 0,
          input: { type: 'formula', source: '=SUM(' },
        },
      ],
      environment,
    );
    expect(invalid.values.get('sheet-1!B1')).toEqual({ type: 'error', value: '#VALUE!' });
    const diagnostics = invalid.diagnostics.get('sheet-1!A1') as unknown[];
    diagnostics.push({ code: 'poison' });
    expect(program.diagnostics.get('sheet-1!A1')).toHaveLength(1);
  });

  it('validates IF arity before selecting a branch', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [
          { row: 0, column: 0, input: { type: 'formula', source: '=IF()' } },
          { row: 0, column: 1, input: { type: 'formula', source: '=IF(TRUE)' } },
        ],
      },
    ]);
    const engine = createFormulaEngine();
    const result = engine.recalculate(engine.compile(document), [], environment);
    expect(result.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#VALUE!' });
    expect(result.values.get('sheet-1!B1')).toEqual({ type: 'error', value: '#VALUE!' });
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

  it('invalidates stable formula caches when the function registry version changes', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=LATE(4)' } }],
      },
    ]);
    const functions = createFormulaFunctionRegistry();
    const engine = createFormulaEngine({ functions });
    const program = engine.compile(document);
    const first = engine.recalculate(program, [], {
      ...environment,
      functionRegistryVersion: functions.version,
    });
    expect(first.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#NAME?' });

    const unregister = functions.register({
      name: 'LATE',
      parameters: { minimum: 1, maximum: 1 },
      returns: 'number',
      volatility: 'stable',
      mode: 'sync',
      evaluate: ([value]) => ({
        type: 'number',
        value: value?.type === 'number' ? value.value * 2 : 0,
      }),
    });
    const registered = engine.recalculate(program, [], {
      ...environment,
      functionRegistryVersion: functions.version,
    });
    expect(registered.values.get('sheet-1!A1')).toEqual({ type: 'number', value: 8 });
    expect(registered.diagnostics.get('sheet-1!A1')).toBeUndefined();

    unregister();
    const removed = engine.recalculate(program, [], {
      ...environment,
      functionRegistryVersion: functions.version,
    });
    expect(removed.values.get('sheet-1!A1')).toEqual({ type: 'error', value: '#NAME?' });
  });

  it('isolates program ownership and validates the declared registry version', () => {
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=NOW()' } }],
      },
    ]);
    const owner = createFormulaEngine();
    const program = owner.compile(document);
    expect(() => createFormulaEngine().recalculate(program, [], environment)).toThrow(
      /another engine/u,
    );
    expect(() =>
      owner.recalculate(program, [], { ...environment, functionRegistryVersion: 'wrong' }),
    ).toThrow(/registry version/u);
  });
});
