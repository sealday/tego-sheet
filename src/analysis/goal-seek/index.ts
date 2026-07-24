import type { SpreadsheetDocument } from '../../document';
import type { DocumentTransactionEnvelope } from '../../document-controller';
import type { SheetId } from '../../core/types/coordinates';
import {
  createFormulaEngine,
  createFormulaFunctionRegistry,
  formulaAddressKey,
  transitiveDependents,
} from '../../formula';
import type {
  CalculationEnvironment,
  FormulaAddress,
  FormulaAst,
  FormulaEngineOptions,
  FormulaFunctionRegistry,
} from '../../formula';

/** Input for one isolated, bounded formula goal-seek run. */
export interface FormulaGoalSeekOptions {
  /** Immutable document snapshot used to compile the isolated calculation program. */
  readonly document: SpreadsheetDocument;
  /** Controller revision from which `document` was captured. */
  readonly sourceRevision: number;
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
  /** Hard wall-clock budget in milliseconds; defaults to 1,000. */
  readonly maximumDurationMs?: number;
  /** Maximum AST/range evaluation steps for one candidate; defaults to 100,000. */
  readonly maximumCalculationSteps?: number;
  /** Injectable monotonic clock used only to enforce the wall-clock budget. */
  readonly budgetClock?: { readonly now: () => number };
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
  readonly status: 'converged' | 'not-converged' | 'invalid-model' | 'out-of-bounds' | 'cancelled';
  /** Best variable value observed, or the initial guess when cancelled before evaluation. */
  readonly value: number;
  /** Target formula result at `value`, or `NaN` when cancelled before evaluation. */
  readonly targetValue: number;
  /** Number of candidate evaluations performed. */
  readonly iterations: number;
  /** Union of formula addresses evaluated for candidates, in stable order. */
  readonly evaluatedAddresses: readonly string[];
  /** Revision to which this suggestion is bound. */
  readonly sourceRevision: number;
  /** Variable address proposed for a later atomic transaction. */
  readonly variable: FormulaAddress;
  /** Stable reason when a bounded run does not converge. */
  readonly diagnostic?: 'GOAL_LIMIT_EXCEEDED' | 'GOAL_OUT_OF_BOUNDS';
}

/** Revision-checked transaction proposal that a host may preview and apply atomically. */
export type FormulaGoalSeekApplyProposal =
  | {
      readonly status: 'ready';
      readonly transaction: DocumentTransactionEnvelope;
    }
  | {
      readonly status: 'rejected';
      readonly code: 'GOAL_RESULT_STALE' | 'GOAL_RESULT_NOT_CONVERGED';
    };

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

function positiveInteger(value: number, name: string, maximum: number): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw new RangeError(`${name} must be an integer from 1 through ${maximum}`);
  }
  return value;
}

function containsDisallowedFunction(
  ast: FormulaAst,
  registry: FormulaFunctionRegistry,
): 'volatile' | 'async' | undefined {
  if (ast.kind === 'call') {
    const definition = registry.resolve(ast.name);
    if (definition?.volatility === 'volatile') return 'volatile';
    if (definition?.mode === 'async') return 'async';
    for (const argument of ast.arguments) {
      const disallowed = containsDisallowedFunction(argument, registry);
      if (disallowed !== undefined) return disallowed;
    }
  } else if (ast.kind === 'binary') {
    return (
      containsDisallowedFunction(ast.left, registry) ??
      containsDisallowedFunction(ast.right, registry)
    );
  } else if (ast.kind === 'unary') {
    return containsDisallowedFunction(ast.operand, registry);
  }
  return undefined;
}

function isolateAffectedDocument(
  document: SpreadsheetDocument,
  affected: ReadonlySet<string>,
): SpreadsheetDocument {
  return {
    ...document,
    workbook: {
      ...document.workbook,
      sheets: document.workbook.sheets.map((sheet) => ({
        ...sheet,
        cells: sheet.cells.filter(
          ({ row, column, cell }) =>
            cell.input.type !== 'formula' ||
            affected.has(formulaAddressKey({ sheetId: sheet.id, row, column })),
        ),
      })),
    },
  };
}

function frozenAddress(address: FormulaAddress): FormulaAddress {
  return Object.freeze({
    sheetId: address.sheetId,
    row: address.row,
    column: address.column,
  });
}

/**
 * Creates a JSON-safe, single-command transaction only while the originating revision is current.
 *
 * This function never receives or invokes a controller.
 */
export function createFormulaGoalSeekApplyProposal(
  result: FormulaGoalSeekResult,
  currentRevision: number,
  transactionId = `goal-seek-${result.sourceRevision}`,
): FormulaGoalSeekApplyProposal {
  if (!Number.isSafeInteger(currentRevision) || currentRevision < 0) {
    throw new RangeError('Goal seek current revision must be a non-negative safe integer');
  }
  if (currentRevision !== result.sourceRevision) {
    return Object.freeze({ status: 'rejected', code: 'GOAL_RESULT_STALE' });
  }
  if (result.status !== 'converged') {
    return Object.freeze({
      status: 'rejected',
      code: 'GOAL_RESULT_NOT_CONVERGED',
    });
  }
  if (!/^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(transactionId)) {
    throw new TypeError('Goal seek transaction ID is invalid');
  }
  return Object.freeze({
    status: 'ready',
    transaction: Object.freeze({
      schemaVersion: 1,
      id: transactionId,
      baseRevision: currentRevision,
      commands: Object.freeze([
        Object.freeze({
          schemaVersion: 1,
          id: `${transactionId}-variable`,
          command: Object.freeze({
            type: 'set-cell-input',
            address: Object.freeze({
              sheet: result.variable.sheetId as SheetId,
              row: result.variable.row,
              column: result.variable.column,
            }),
            input: Object.freeze({ type: 'number', value: result.value }),
          }),
        }),
      ]),
    }),
  });
}

/**
 * Finds a numeric variable value with a secant search in a private formula program.
 *
 * The source document and any controller remain untouched. After the private program's initial
 * compile, each candidate invalidates only the variable cell's transitive formula dependents.
 */
export function solveFormulaGoalSeek(options: FormulaGoalSeekOptions): FormulaGoalSeekResult {
  const cancelled = (): boolean => options.signal?.aborted ?? false;
  if (!Number.isSafeInteger(options.sourceRevision) || options.sourceRevision < 0) {
    throw new RangeError('Goal seek sourceRevision must be a non-negative safe integer');
  }
  if (
    options.engine?.functions !== undefined ||
    options.engine?.names !== undefined ||
    options.engine?.tables !== undefined
  ) {
    throw new TypeError('Goal seek forbids custom functions and external resolvers');
  }
  const targetValue = finite(options.targetValue, 'Goal seek target');
  const tolerance = finite(options.tolerance ?? 1e-7, 'Goal seek tolerance');
  if (tolerance <= 0) throw new RangeError('Goal seek tolerance must be positive');
  const maximumIterations = positiveInteger(
    options.maximumIterations ?? 100,
    'Goal seek maximumIterations',
    1_000,
  );
  const maximumCalculationSteps = positiveInteger(
    options.maximumCalculationSteps ?? options.engine?.maximumCalculationSteps ?? 100_000,
    'Goal seek maximumCalculationSteps',
    10_000_000,
  );
  const maximumDurationMs = finite(
    options.maximumDurationMs ?? 1_000,
    'Goal seek maximumDurationMs',
  );
  if (maximumDurationMs <= 0) {
    throw new RangeError('Goal seek maximumDurationMs must be positive');
  }
  const budgetClock = options.budgetClock ?? {
    now: () => globalThis.performance?.now() ?? Date.now(),
  };
  const startedAt = budgetClock.now();
  const durationExceeded = (): boolean => budgetClock.now() - startedAt > maximumDurationMs;
  const minimum = finite(options.minimum ?? -Number.MAX_VALUE, 'Goal seek minimum');
  const maximum = finite(options.maximum ?? Number.MAX_VALUE, 'Goal seek maximum');
  if (minimum > maximum) throw new RangeError('Goal seek minimum must not exceed maximum');
  const currentValue = numericVariable(options.document, options.variable);
  const requestedInitialGuess = finite(
    options.initialGuess ?? currentValue,
    'Goal seek initialGuess',
  );
  const initialGuess = bounded(requestedInitialGuess, minimum, maximum);
  const resultBase = {
    sourceRevision: options.sourceRevision,
    variable: frozenAddress(options.variable),
  } as const;
  if (requestedInitialGuess < minimum || requestedInitialGuess > maximum) {
    return Object.freeze({
      ...resultBase,
      status: 'out-of-bounds',
      value: initialGuess,
      targetValue: Number.NaN,
      iterations: 0,
      evaluatedAddresses: Object.freeze([]),
      diagnostic: 'GOAL_OUT_OF_BOUNDS',
    });
  }
  if (cancelled()) {
    return Object.freeze({
      ...resultBase,
      status: 'cancelled',
      value: initialGuess,
      targetValue: Number.NaN,
      iterations: 0,
      evaluatedAddresses: Object.freeze([]),
    });
  }

  const compiler = createFormulaEngine({
    maximumEvaluations: options.engine?.maximumEvaluations,
    maximumCalculationSteps,
    maximumSpillCells: options.engine?.maximumSpillCells,
  });
  const compiledDocument = compiler.compile(options.document);
  const variableKey = formulaAddressKey(options.variable);
  const targetKey = formulaAddressKey(options.target);
  const affected = transitiveDependents(compiledDocument.graph, [variableKey]);
  if (!affected.has(targetKey)) {
    throw new TypeError('Goal seek target does not depend on the variable cell');
  }
  const registry = createFormulaFunctionRegistry();
  for (const address of affected) {
    const ast = compiledDocument.formulas.get(address);
    if (ast === undefined) continue;
    const disallowed = containsDisallowedFunction(ast, registry);
    if (disallowed !== undefined) {
      throw new TypeError(`Goal seek forbids ${disallowed} formula functions`);
    }
  }

  const engine = createFormulaEngine({
    maximumEvaluations: options.engine?.maximumEvaluations,
    maximumCalculationSteps,
    maximumSpillCells: options.engine?.maximumSpillCells,
  });
  const program = engine.compile(isolateAffectedDocument(options.document, affected));
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
      ...resultBase,
      status,
      value: best.x,
      targetValue: best.y,
      iterations,
      evaluatedAddresses: Object.freeze([...evaluated].sort()),
    });

  const limited = (
    best: { x: number; y: number } = { x: initialGuess, y: Number.NaN },
  ): FormulaGoalSeekResult =>
    Object.freeze({
      ...finish('not-converged', best),
      diagnostic: 'GOAL_LIMIT_EXCEEDED',
    });

  if (durationExceeded()) return limited();
  let previous = evaluate(initialGuess);
  let best = previous;
  if (Math.abs(previous.residual) <= tolerance) return finish('converged', best);
  if (iterations >= maximumIterations || durationExceeded()) return limited(best);

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
    if (durationExceeded()) return limited(best);
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
  return limited(best);
}
