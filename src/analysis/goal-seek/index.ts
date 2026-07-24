import type { SpreadsheetDocument } from '../../document';
import { createFormulaEngine, formulaAddressKey, transitiveDependents } from '../../formula';
import type { CalculationEnvironment, FormulaAddress, FormulaEngineOptions } from '../../formula';

/** Input for one isolated, bounded formula goal-seek run. */
export interface FormulaGoalSeekOptions {
  /** Immutable document snapshot used to compile the isolated calculation program. */
  readonly document: SpreadsheetDocument;
  /** Numeric input cell adjusted by the solver. */
  readonly variable: FormulaAddress;
  /** Formula cell whose numeric result should reach `targetValue`. */
  readonly target: FormulaAddress;
  /** Desired finite numeric result. */
  readonly targetValue: number;
  /** First candidate; defaults to the variable cell's current numeric input. */
  readonly initialGuess?: number;
  /** Finite non-zero distance to the second secant candidate. */
  readonly initialStep?: number;
  /** Optional inclusive lower bound for candidate values. */
  readonly minimum?: number;
  /** Optional inclusive upper bound for candidate values. */
  readonly maximum?: number;
  /** Absolute result tolerance; defaults to `1e-7`. */
  readonly tolerance?: number;
  /** Hard iteration budget, from 1 through 1,000; defaults to 100. */
  readonly maximumIterations?: number;
  /** Optional cooperative cancellation signal. */
  readonly signal?: AbortSignal;
  /** Explicit deterministic calculation environment. */
  readonly environment: CalculationEnvironment;
  /** Optional formula-engine limits and registries. */
  readonly engine?: FormulaEngineOptions;
}

/** Stable outcome of a bounded formula goal-seek run. */
export interface FormulaGoalSeekResult {
  /** Completion status. */
  readonly status: 'converged' | 'iteration-limit' | 'cancelled';
  /** Best variable value observed, or the initial guess when cancelled before evaluation. */
  readonly value: number;
  /** Target formula result at `value`, or `NaN` when cancelled before evaluation. */
  readonly targetValue: number;
  /** Number of candidate evaluations performed. */
  readonly iterations: number;
  /** Union of formula addresses evaluated for candidates, in stable order. */
  readonly evaluatedAddresses: readonly string[];
}

function numericVariable(document: SpreadsheetDocument, address: FormulaAddress): number {
  const sheet = document.workbook.sheets.find(({ id }) => id === address.sheetId);
  const input = sheet?.cells.find(
    ({ row, column }) => row === address.row && column === address.column,
  )?.cell.input;
  if (input?.type !== 'number') {
    throw new TypeError('Goal seek variable must contain a finite numeric input');
  }
  return input.value;
}

function finite(value: number, name: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${name} must be finite`);
  return value;
}

function bounded(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

/**
 * Finds a numeric variable value with a secant search in a private formula program.
 *
 * The source document and any controller remain untouched. After the private program's initial
 * compile, each candidate invalidates only the variable cell's transitive formula dependents.
 */
export function solveFormulaGoalSeek(options: FormulaGoalSeekOptions): FormulaGoalSeekResult {
  const cancelled = (): boolean => options.signal?.aborted ?? false;
  const targetValue = finite(options.targetValue, 'Goal seek target');
  const tolerance = finite(options.tolerance ?? 1e-7, 'Goal seek tolerance');
  if (tolerance <= 0) throw new RangeError('Goal seek tolerance must be positive');
  const maximumIterations = options.maximumIterations ?? 100;
  if (
    !Number.isSafeInteger(maximumIterations) ||
    maximumIterations < 1 ||
    maximumIterations > 1_000
  ) {
    throw new RangeError('Goal seek maximumIterations must be an integer from 1 through 1000');
  }
  const minimum = finite(options.minimum ?? -Number.MAX_VALUE, 'Goal seek minimum');
  const maximum = finite(options.maximum ?? Number.MAX_VALUE, 'Goal seek maximum');
  if (minimum > maximum) throw new RangeError('Goal seek minimum must not exceed maximum');
  const currentValue = numericVariable(options.document, options.variable);
  const initialGuess = bounded(
    finite(options.initialGuess ?? currentValue, 'Goal seek initialGuess'),
    minimum,
    maximum,
  );
  if (cancelled()) {
    return Object.freeze({
      status: 'cancelled',
      value: initialGuess,
      targetValue: Number.NaN,
      iterations: 0,
      evaluatedAddresses: Object.freeze([]),
    });
  }

  const engine = createFormulaEngine(options.engine);
  const program = engine.compile(options.document);
  const variableKey = formulaAddressKey(options.variable);
  const targetKey = formulaAddressKey(options.target);
  if (!transitiveDependents(program.graph, [variableKey]).has(targetKey)) {
    throw new TypeError('Goal seek target does not depend on the variable cell');
  }

  engine.recalculate(program, [], options.environment);
  const evaluated = new Set<string>();
  let iterations = 0;
  const evaluate = (candidate: number): { x: number; y: number; residual: number } => {
    const x = bounded(candidate, minimum, maximum);
    const calculation = engine.recalculate(
      program,
      [{ ...options.variable, input: { type: 'number', value: x } }],
      options.environment,
    );
    iterations += 1;
    for (const address of calculation.evaluatedAddresses) evaluated.add(address);
    const result = calculation.values.get(targetKey);
    if (result?.type !== 'number' || !Number.isFinite(result.value)) {
      throw new TypeError('Goal seek target formula must produce a finite numeric result');
    }
    return { x, y: result.value, residual: result.value - targetValue };
  };
  const finish = (
    status: FormulaGoalSeekResult['status'],
    best: { x: number; y: number },
  ): FormulaGoalSeekResult =>
    Object.freeze({
      status,
      value: best.x,
      targetValue: best.y,
      iterations,
      evaluatedAddresses: Object.freeze([...evaluated].sort()),
    });

  let previous = evaluate(initialGuess);
  let best = previous;
  if (Math.abs(previous.residual) <= tolerance) return finish('converged', best);
  if (iterations >= maximumIterations) return finish('iteration-limit', best);

  const initialStep = finite(
    options.initialStep ?? Math.max(1, Math.abs(initialGuess) * 0.1),
    'Goal seek initialStep',
  );
  if (initialStep === 0) throw new RangeError('Goal seek initialStep must not be zero');
  let current = evaluate(bounded(initialGuess + initialStep, minimum, maximum));
  if (Math.abs(current.residual) < Math.abs(best.residual)) best = current;
  if (Math.abs(current.residual) <= tolerance) return finish('converged', best);

  while (iterations < maximumIterations) {
    if (cancelled()) return finish('cancelled', best);
    const denominator = current.residual - previous.residual;
    let candidate =
      denominator === 0
        ? current.x + initialStep * (iterations + 1)
        : current.x - (current.residual * (current.x - previous.x)) / denominator;
    candidate = bounded(candidate, minimum, maximum);
    if (candidate === current.x) {
      const direction = current.residual > 0 ? -1 : 1;
      candidate = bounded(
        current.x + direction * Math.abs(initialStep) * (iterations + 1),
        minimum,
        maximum,
      );
    }
    if (candidate === current.x) break;
    previous = current;
    current = evaluate(candidate);
    if (Math.abs(current.residual) < Math.abs(best.residual)) best = current;
    if (Math.abs(current.residual) <= tolerance) return finish('converged', best);
  }
  return finish('iteration-limit', best);
}
