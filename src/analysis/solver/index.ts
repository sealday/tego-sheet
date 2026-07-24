import type { FormulaAddress } from '../../formula';
import type { AdapterRegistry } from '../../sdk/adapters';
import { createCapabilityGrant } from '../../sdk/trust';

/** Optimization direction for a solver objective. */
export type SolverGoal = 'minimize' | 'maximize' | 'target';

/** One bounded decision variable in the solver problem IR. */
export interface SolverVariable {
  readonly id: string;
  readonly address: FormulaAddress;
  readonly minimum?: number;
  readonly maximum?: number;
  readonly integer?: boolean;
}

/** One scalar formula-cell constraint in the solver problem IR. */
export interface SolverConstraint {
  readonly id: string;
  readonly address: FormulaAddress;
  readonly operator: '<=' | '=' | '>=';
  readonly value: number;
}

/** Host request compiled into transport-safe solver input. */
export interface SolverModelRequest {
  readonly documentId: string;
  readonly revision: string;
  readonly objective: {
    readonly goal: SolverGoal;
    readonly address: FormulaAddress;
    readonly targetValue?: number;
  };
  readonly variables: readonly SolverVariable[];
  readonly constraints: readonly SolverConstraint[];
}

/** Immutable, JSON-safe problem sent across the isolated adapter boundary. */
export type CompiledSolverModel = SolverModelRequest;

/** Stable solver completion state. */
export type SolverStatus =
  | 'optimal'
  | 'feasible'
  | 'infeasible'
  | 'unbounded'
  | 'timeout'
  | 'cancelled'
  | 'error';

/** Validated result returned by a solver adapter. */
export interface SolverResult {
  readonly status: SolverStatus;
  readonly objectiveValue?: number;
  readonly candidates: readonly {
    readonly variableId: string;
    readonly value: number;
  }[];
  readonly residuals: readonly {
    readonly constraintId: string;
    readonly value: number;
  }[];
}

export interface RunSolverOptions {
  readonly signal: AbortSignal;
  readonly adapterId?: string;
  readonly maximumDurationMs?: number;
  readonly maximumInputBytes?: number;
  readonly maximumOutputBytes?: number;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/;

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function finite(value: number, label: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${label} must be finite`);
  return value;
}

function address(value: FormulaAddress, label: string): FormulaAddress {
  if (
    value === null ||
    typeof value !== 'object' ||
    !identifierPattern.test(value.sheetId) ||
    !Number.isSafeInteger(value.row) ||
    value.row < 0 ||
    value.row > 1_048_575 ||
    !Number.isSafeInteger(value.column) ||
    value.column < 0 ||
    value.column > 16_383
  ) {
    throw new TypeError(`${label} address is invalid`);
  }
  return Object.freeze({
    sheetId: value.sheetId,
    row: value.row,
    column: value.column,
  });
}

/** Validates and snapshots a solver model without retaining caller objects. */
export function compileSolverModel(request: SolverModelRequest): CompiledSolverModel {
  if (request === null || typeof request !== 'object') {
    throw new TypeError('Solver request must be an object');
  }
  if (request.variables.length === 0 || request.variables.length > 10_000) {
    throw new RangeError('Solver variable limit is from 1 through 10000');
  }
  if (request.constraints.length > 100_000) {
    throw new RangeError('Solver constraint limit is 100000');
  }
  const variableIds = new Set<string>();
  const variables = request.variables.map((variable, index): SolverVariable => {
    const id = identifier(variable.id, `Solver variable ${index} ID`);
    if (variableIds.has(id)) throw new TypeError(`Duplicate solver variable ID ${id}`);
    variableIds.add(id);
    const minimum =
      variable.minimum === undefined
        ? undefined
        : finite(variable.minimum, `Solver variable ${id} minimum`);
    const maximum =
      variable.maximum === undefined
        ? undefined
        : finite(variable.maximum, `Solver variable ${id} maximum`);
    if (minimum !== undefined && maximum !== undefined && minimum > maximum) {
      throw new RangeError(`Solver variable ${id} bounds are invalid`);
    }
    return Object.freeze({
      id,
      address: address(variable.address, `Solver variable ${id}`),
      ...(minimum === undefined ? {} : { minimum }),
      ...(maximum === undefined ? {} : { maximum }),
      ...(variable.integer === undefined ? {} : { integer: variable.integer === true }),
    });
  });
  const constraintIds = new Set<string>();
  const constraints = request.constraints.map((constraint, index): SolverConstraint => {
    const id = identifier(constraint.id, `Solver constraint ${index} ID`);
    if (constraintIds.has(id)) throw new TypeError(`Duplicate solver constraint ID ${id}`);
    constraintIds.add(id);
    if (!['<=', '=', '>='].includes(constraint.operator)) {
      throw new TypeError(`Solver constraint ${id} operator is invalid`);
    }
    return Object.freeze({
      id,
      address: address(constraint.address, `Solver constraint ${id}`),
      operator: constraint.operator,
      value: finite(constraint.value, `Solver constraint ${id} value`),
    });
  });
  if (!['minimize', 'maximize', 'target'].includes(request.objective.goal)) {
    throw new TypeError('Solver objective goal is invalid');
  }
  const targetValue =
    request.objective.targetValue === undefined
      ? undefined
      : finite(request.objective.targetValue, 'Solver objective targetValue');
  if (request.objective.goal === 'target' && targetValue === undefined) {
    throw new TypeError('Target solver objectives require targetValue');
  }
  return Object.freeze({
    documentId: identifier(request.documentId, 'Solver documentId'),
    revision: identifier(request.revision, 'Solver revision'),
    objective: Object.freeze({
      goal: request.objective.goal,
      address: address(request.objective.address, 'Solver objective'),
      ...(targetValue === undefined ? {} : { targetValue }),
    }),
    variables: Object.freeze(variables),
    constraints: Object.freeze(constraints),
  });
}

function isSolverResult(value: unknown, model: CompiledSolverModel): value is SolverResult {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const result = value as Partial<SolverResult>;
  if (
    !['optimal', 'feasible', 'infeasible', 'unbounded', 'timeout', 'cancelled', 'error'].includes(
      result.status ?? '',
    ) ||
    (result.objectiveValue !== undefined && !Number.isFinite(result.objectiveValue)) ||
    !Array.isArray(result.candidates) ||
    !Array.isArray(result.residuals)
  ) {
    return false;
  }
  const variables = new Set(model.variables.map(({ id }) => id));
  const candidateIds = new Set<string>();
  for (const candidate of result.candidates) {
    if (
      candidate === null ||
      typeof candidate !== 'object' ||
      !variables.has(candidate.variableId) ||
      candidateIds.has(candidate.variableId) ||
      !Number.isFinite(candidate.value)
    ) {
      return false;
    }
    candidateIds.add(candidate.variableId);
  }
  const constraints = new Set(model.constraints.map(({ id }) => id));
  const residualIds = new Set<string>();
  for (const residual of result.residuals) {
    if (
      residual === null ||
      typeof residual !== 'object' ||
      !constraints.has(residual.constraintId) ||
      residualIds.has(residual.constraintId) ||
      !Number.isFinite(residual.value)
    ) {
      return false;
    }
    residualIds.add(residual.constraintId);
  }
  return true;
}

/** Executes a compiled model through the public isolated-worker solver adapter contract. */
export async function runSolver(
  registry: AdapterRegistry,
  model: CompiledSolverModel,
  options: RunSolverOptions,
): Promise<SolverResult> {
  const resolution = registry.resolve('solver', {
    ...(options.adapterId === undefined ? {} : { id: options.adapterId }),
    capability: 'solve',
  });
  if (resolution.manifest.execution !== 'isolated-worker') {
    throw new TypeError('Solver adapters must execute as isolated-worker extensions');
  }
  const scope = registry.createScope({
    documentId: model.documentId,
    signal: options.signal,
    grant: createCapabilityGrant(['solve']),
    limits: {
      maxConcurrentInvocations: 1,
      maxDurationMs: options.maximumDurationMs ?? 30_000,
      maxInputBytes: options.maximumInputBytes ?? 1024 * 1024,
      maxOutputBytes: options.maximumOutputBytes ?? 4 * 1024 * 1024,
    },
  });
  try {
    return await scope.invoke(resolution, {
      capability: 'solve',
      input: model,
      validateResult: (value): value is SolverResult => isSolverResult(value, model),
    });
  } finally {
    await scope.dispose();
  }
}
