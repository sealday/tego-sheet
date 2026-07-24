/** Scalar value accepted by the deterministic Pivot input snapshot. */
export type PivotScalar = string | number | boolean | null;

/** Immutable tabular source captured at one revision. */
export interface PivotSourceSnapshot {
  readonly revision: string;
  readonly fields: readonly string[];
  readonly rows: readonly (readonly PivotScalar[])[];
}

/** Supported built-in Pivot aggregation. */
export type PivotAggregate = 'sum' | 'count' | 'average' | 'min' | 'max';

/** Persistent renderer-neutral Pivot definition. */
export interface PivotDefinition {
  readonly id: string;
  readonly rows: readonly string[];
  readonly columns: readonly string[];
  readonly values: readonly {
    readonly id: string;
    readonly field: string;
    readonly aggregate: PivotAggregate;
  }[];
  readonly filters: readonly {
    readonly field: string;
    readonly values: readonly PivotScalar[];
  }[];
}

/** One non-empty aggregated Pivot output cell. */
export interface PivotResultCell {
  readonly rowKey: readonly PivotScalar[];
  readonly columnKey: readonly PivotScalar[];
  readonly values: Readonly<Record<string, number>>;
  readonly sourceRows: readonly number[];
}

/** Last successful immutable Pivot output. */
export interface PivotResult {
  readonly definitionId: string;
  readonly sourceRevision: string;
  readonly rowKeys: readonly (readonly PivotScalar[])[];
  readonly columnKeys: readonly (readonly PivotScalar[])[];
  readonly cells: readonly PivotResultCell[];
}

/** Hard limits for one Pivot refresh. */
export interface PivotRefreshLimits {
  readonly maximumRows: number;
  readonly maximumFields: number;
  readonly maximumCardinality: number;
  readonly maximumResultCells: number;
  /** Comparisons and materialized values allowed after aggregation. */
  readonly maximumPostProcessingSteps: number;
}

/** Options for a cancellable Pivot refresh. */
export interface PivotRefreshOptions {
  readonly signal: AbortSignal;
  readonly previous?: PivotResult;
  readonly limits?: Partial<PivotRefreshLimits>;
  readonly onProgress?: (completedRows: number, totalRows: number) => void;
}

/** Atomic refresh outcome; cancellation may retain the last successful result. */
export type PivotRefreshOutcome =
  | {
      readonly status: 'ready';
      readonly stale: false;
      readonly result: PivotResult;
    }
  | {
      readonly status: 'cancelled';
      readonly stale: true;
      readonly result?: PivotResult;
    };

const defaultLimits: Readonly<PivotRefreshLimits> = Object.freeze({
  maximumRows: 100_000,
  maximumFields: 1_000,
  maximumCardinality: 10_000,
  maximumResultCells: 100_000,
  maximumPostProcessingSteps: 10_000_000,
});
const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/;

interface AggregateState {
  sum: number;
  count: number;
  numericCount: number;
  minimum?: number;
  maximum?: number;
}

interface GroupState {
  readonly rowKey: readonly PivotScalar[];
  readonly columnKey: readonly PivotScalar[];
  readonly values: Map<string, AggregateState>;
  readonly sourceRows: number[];
}

function cancelled(previous: PivotResult | undefined): PivotRefreshOutcome {
  return Object.freeze({
    status: 'cancelled',
    stale: true,
    ...(previous === undefined ? {} : { result: previous }),
  });
}

function matchingPrevious(
  definitionId: string,
  previous: PivotResult | undefined,
): PivotResult | undefined {
  return previous?.definitionId === definitionId ? previous : undefined;
}

function checkedLimit(value: number | undefined, fallback: number, label: string): number {
  const result = value ?? fallback;
  if (!Number.isSafeInteger(result) || result < 1) {
    throw new RangeError(`${label} must be a positive safe integer`);
  }
  return result;
}

function compareScalar(left: PivotScalar, right: PivotScalar): number {
  const order = (value: PivotScalar): number =>
    value === null ? 0 : typeof value === 'boolean' ? 1 : typeof value === 'number' ? 2 : 3;
  const typeDifference = order(left) - order(right);
  if (typeDifference !== 0) return typeDifference;
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  const leftText = String(left);
  const rightText = String(right);
  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
}

function compareKey(left: readonly PivotScalar[], right: readonly PivotScalar[]): number {
  for (let index = 0; index < Math.max(left.length, right.length); index += 1) {
    if (index >= left.length) return -1;
    if (index >= right.length) return 1;
    const comparison = compareScalar(left[index]!, right[index]!);
    if (comparison !== 0) return comparison;
  }
  return 0;
}

function snapshotKey(values: readonly PivotScalar[]): readonly PivotScalar[] {
  return Object.freeze([...values]);
}

function keyId(values: readonly PivotScalar[]): string {
  return JSON.stringify(values);
}

function validateScalar(value: PivotScalar, label: string): void {
  if (
    value !== null &&
    typeof value !== 'string' &&
    typeof value !== 'boolean' &&
    (typeof value !== 'number' || !Number.isFinite(value))
  ) {
    throw new TypeError(`${label} must be a finite Pivot scalar`);
  }
}

function finalizeAggregate(state: AggregateState, aggregate: PivotAggregate): number {
  if (aggregate === 'count') return state.count;
  if (aggregate === 'sum') return state.sum;
  if (aggregate === 'average') {
    return state.numericCount === 0 ? 0 : state.sum / state.numericCount;
  }
  if (aggregate === 'min') return state.minimum ?? 0;
  return state.maximum ?? 0;
}

interface PostProcessingBudget {
  readonly signal: AbortSignal;
  readonly maximum: number;
  steps: number;
}

async function postProcessingStep(budget: PostProcessingBudget): Promise<boolean> {
  budget.steps += 1;
  if (budget.steps > budget.maximum) {
    throw new RangeError('Pivot post-processing limit exceeded');
  }
  if (budget.steps % 256 === 0) await Promise.resolve();
  return !budget.signal.aborted;
}

async function stableBudgetedSort<Value>(
  values: readonly Value[],
  compare: (left: Value, right: Value) => number,
  budget: PostProcessingBudget,
): Promise<Value[] | undefined> {
  let source = [...values];
  let target = Array.from<Value>({ length: source.length });
  for (let width = 1; width < source.length; width *= 2) {
    for (let start = 0; start < source.length; start += width * 2) {
      let left = start;
      let right = Math.min(start + width, source.length);
      const leftEnd = right;
      const rightEnd = Math.min(start + width * 2, source.length);
      let output = start;
      while (left < leftEnd || right < rightEnd) {
        if (!(await postProcessingStep(budget))) return undefined;
        if (
          right >= rightEnd ||
          (left < leftEnd && compare(source[left] as Value, source[right] as Value) <= 0)
        ) {
          target[output] = source[left] as Value;
          left += 1;
        } else {
          target[output] = source[right] as Value;
          right += 1;
        }
        output += 1;
      }
    }
    [source, target] = [target, source];
  }
  return source;
}

/** Aggregates a source snapshot without mutating the workbook or the previous result. */
export async function refreshPivot(
  source: PivotSourceSnapshot,
  definition: PivotDefinition,
  options: PivotRefreshOptions,
): Promise<PivotRefreshOutcome> {
  const previous = matchingPrevious(definition.id, options.previous);
  if (options.signal.aborted) return cancelled(previous);
  const limits: PivotRefreshLimits = {
    maximumRows: checkedLimit(
      options.limits?.maximumRows,
      defaultLimits.maximumRows,
      'Pivot row limit',
    ),
    maximumFields: checkedLimit(
      options.limits?.maximumFields,
      defaultLimits.maximumFields,
      'Pivot field limit',
    ),
    maximumCardinality: checkedLimit(
      options.limits?.maximumCardinality,
      defaultLimits.maximumCardinality,
      'Pivot cardinality limit',
    ),
    maximumResultCells: checkedLimit(
      options.limits?.maximumResultCells,
      defaultLimits.maximumResultCells,
      'Pivot result cell limit',
    ),
    maximumPostProcessingSteps: checkedLimit(
      options.limits?.maximumPostProcessingSteps,
      defaultLimits.maximumPostProcessingSteps,
      'Pivot post-processing limit',
    ),
  };
  if (source.rows.length > limits.maximumRows) {
    throw new RangeError(`Pivot row limit exceeded: ${source.rows.length}`);
  }
  if (source.fields.length > limits.maximumFields) {
    throw new RangeError(`Pivot field limit exceeded: ${source.fields.length}`);
  }
  if (!identifierPattern.test(definition.id)) throw new TypeError('Pivot ID is invalid');
  const fieldIndexes = new Map<string, number>();
  source.fields.forEach((field, index) => {
    if (!identifierPattern.test(field)) throw new TypeError(`Pivot field ${index} is invalid`);
    if (fieldIndexes.has(field)) throw new TypeError(`Duplicate Pivot field ${field}`);
    fieldIndexes.set(field, index);
  });
  const indexOf = (field: string): number => {
    const index = fieldIndexes.get(field);
    if (index === undefined) throw new TypeError(`Unknown Pivot field ${field}`);
    return index;
  };
  const rowIndexes = definition.rows.map(indexOf);
  const columnIndexes = definition.columns.map(indexOf);
  const valueDefinitions = definition.values.map((value) => ({
    ...value,
    fieldIndex: indexOf(value.field),
  }));
  if (valueDefinitions.length === 0) {
    throw new TypeError('Pivot requires at least one value field');
  }
  const valueIds = new Set<string>();
  for (const value of valueDefinitions) {
    if (!identifierPattern.test(value.id) || valueIds.has(value.id)) {
      throw new TypeError(`Pivot value ID ${value.id} is invalid or duplicated`);
    }
    if (!['sum', 'count', 'average', 'min', 'max'].includes(value.aggregate)) {
      throw new TypeError(`Pivot aggregate ${String(value.aggregate)} is invalid`);
    }
    valueIds.add(value.id);
  }
  const filters = definition.filters.map((filter) => ({
    fieldIndex: indexOf(filter.field),
    values: new Set(filter.values.map((value) => keyId([value]))),
  }));
  const groups = new Map<string, GroupState>();
  const rowKeys = new Map<string, readonly PivotScalar[]>();
  const columnKeys = new Map<string, readonly PivotScalar[]>();
  for (let sourceRow = 0; sourceRow < source.rows.length; sourceRow += 1) {
    if (options.signal.aborted) return cancelled(previous);
    const row = source.rows[sourceRow]!;
    if (row.length !== source.fields.length) {
      throw new TypeError(`Pivot source row ${sourceRow} has an invalid field count`);
    }
    row.forEach((value, index) => validateScalar(value, `Pivot row ${sourceRow} field ${index}`));
    if (
      filters.some(
        (filter) =>
          filter.values.size > 0 && !filter.values.has(keyId([row[filter.fieldIndex] ?? null])),
      )
    ) {
      continue;
    }
    const rowKey = snapshotKey(rowIndexes.map((index) => row[index] ?? null));
    const columnKey = snapshotKey(columnIndexes.map((index) => row[index] ?? null));
    const rowId = keyId(rowKey);
    const columnId = keyId(columnKey);
    rowKeys.set(rowId, rowKey);
    columnKeys.set(columnId, columnKey);
    if (rowKeys.size > limits.maximumCardinality || columnKeys.size > limits.maximumCardinality) {
      throw new RangeError('Pivot cardinality limit exceeded');
    }
    const groupId = `${rowId}\u0000${columnId}`;
    let group = groups.get(groupId);
    if (group === undefined) {
      if (groups.size >= limits.maximumResultCells) {
        throw new RangeError('Pivot result cell limit exceeded');
      }
      group = { rowKey, columnKey, values: new Map(), sourceRows: [] };
      groups.set(groupId, group);
    }
    group.sourceRows.push(sourceRow);
    for (const value of valueDefinitions) {
      const current = group.values.get(value.id) ?? {
        sum: 0,
        count: 0,
        numericCount: 0,
      };
      const cell = row[value.fieldIndex] ?? null;
      if (cell !== null) current.count += 1;
      if (typeof cell === 'number') {
        current.sum += cell;
        current.numericCount += 1;
        current.minimum = current.minimum === undefined ? cell : Math.min(current.minimum, cell);
        current.maximum = current.maximum === undefined ? cell : Math.max(current.maximum, cell);
      }
      group.values.set(value.id, current);
    }
    if ((sourceRow + 1) % 512 === 0) {
      options.onProgress?.(sourceRow + 1, source.rows.length);
      await Promise.resolve();
    }
  }
  options.onProgress?.(source.rows.length, source.rows.length);
  if (options.signal.aborted) return cancelled(previous);
  const budget: PostProcessingBudget = {
    signal: options.signal,
    maximum: limits.maximumPostProcessingSteps,
    steps: 0,
  };
  const sortedGroups = await stableBudgetedSort(
    [...groups.values()],
    (left, right) =>
      compareKey(left.rowKey, right.rowKey) || compareKey(left.columnKey, right.columnKey),
    budget,
  );
  if (sortedGroups === undefined) return cancelled(previous);
  const cells: PivotResultCell[] = [];
  for (const group of sortedGroups) {
    const values: Record<string, number> = {};
    for (const definition_ of valueDefinitions) {
      if (!(await postProcessingStep(budget))) return cancelled(previous);
      values[definition_.id] = finalizeAggregate(
        group.values.get(definition_.id) ?? {
          sum: 0,
          count: 0,
          numericCount: 0,
        },
        definition_.aggregate,
      );
    }
    if (!(await postProcessingStep(budget))) return cancelled(previous);
    cells.push(
      Object.freeze({
        rowKey: group.rowKey,
        columnKey: group.columnKey,
        values: Object.freeze(values),
        sourceRows: Object.freeze([...group.sourceRows]),
      }),
    );
  }
  const sortedRowKeys = await stableBudgetedSort([...rowKeys.values()], compareKey, budget);
  if (sortedRowKeys === undefined) return cancelled(previous);
  const sortedColumnKeys = await stableBudgetedSort([...columnKeys.values()], compareKey, budget);
  if (sortedColumnKeys === undefined) return cancelled(previous);
  const result: PivotResult = Object.freeze({
    definitionId: definition.id,
    sourceRevision: source.revision,
    rowKeys: Object.freeze(sortedRowKeys),
    columnKeys: Object.freeze(sortedColumnKeys),
    cells: Object.freeze(cells),
  });
  return Object.freeze({ status: 'ready', stale: false, result });
}
