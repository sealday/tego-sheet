import type { PivotScalar } from '../pivot';

/** Persistent or session-owned Slicer definition. */
export interface SlicerDefinition {
  readonly id: string;
  readonly fieldId: string;
  readonly targets: readonly string[];
  readonly selection: readonly PivotScalar[];
  readonly stateScope: 'document' | 'session';
}

/** Stable recoverable Slicer compilation diagnostic. */
export interface SlicerDiagnostic {
  readonly code: 'SLICER_TARGET_MISSING';
  readonly slicerId: string;
  readonly targetId: string;
  readonly message: string;
}

/** Session-only selections keyed by Slicer ID. */
export type SlicerSessionSelections = Readonly<Record<string, readonly PivotScalar[]>>;

/** Shared target-aware predicate context consumed by Tables, Pivot, and Charts. */
export interface SlicerFilterContext {
  readonly diagnostics: readonly SlicerDiagnostic[];
  matches(targetId: string, fields: Readonly<Record<string, PivotScalar>>): boolean;
}

export interface CompileSlicerOptions {
  readonly knownTargets?: readonly string[];
  readonly maximumSlicers?: number;
  readonly maximumSelections?: number;
  readonly maximumTargets?: number;
}

/** One unique value and its source occurrence count. */
export interface SlicerValueIndexEntry {
  readonly value: PivotScalar;
  readonly count: number;
}

export interface SlicerValueIndexOptions {
  readonly maximumValues?: number;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/;

function scalarKey(value: PivotScalar): string {
  if (typeof value === 'number' && !Number.isFinite(value)) {
    throw new TypeError('Slicer values must be finite');
  }
  return `${value === null ? 'null' : typeof value}:${JSON.stringify(value)}`;
}

function limit(value: number | undefined, fallback: number, label: string): number {
  const resolved = value ?? fallback;
  if (!Number.isSafeInteger(resolved) || resolved < 1) {
    throw new RangeError(`${label} must be a positive safe integer`);
  }
  return resolved;
}

function scalarOrder(value: PivotScalar): number {
  return value === null ? 0 : typeof value === 'boolean' ? 1 : typeof value === 'number' ? 2 : 3;
}

function compareScalar(left: PivotScalar, right: PivotScalar): number {
  const type = scalarOrder(left) - scalarOrder(right);
  if (type !== 0) return type;
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  const leftText = String(left);
  const rightText = String(right);
  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
}

/**
 * Compiles Slicer selections into a shared predicate context.
 *
 * Each Slicer's selected values are OR-ed. Every Slicer contributing to the same target is AND-ed.
 * Empty selections contribute no predicate. Session state is read from the supplied overlay and is
 * never written back to persistent definitions.
 */
export function compileSlicerFilterContext(
  slicers: readonly SlicerDefinition[],
  sessionSelections: SlicerSessionSelections = {},
  options: CompileSlicerOptions = {},
): SlicerFilterContext {
  const maximumSlicers = limit(options.maximumSlicers, 10_000, 'Slicer limit');
  const maximumSelections = limit(options.maximumSelections, 10_000, 'Slicer selection limit');
  const maximumTargets = limit(options.maximumTargets, 10_000, 'Slicer target limit');
  if (slicers.length > maximumSlicers) throw new RangeError('Slicer limit exceeded');
  const knownTargets =
    options.knownTargets === undefined ? undefined : new Set(options.knownTargets);
  const diagnostics: SlicerDiagnostic[] = [];
  const predicates = new Map<
    string,
    readonly { readonly fieldId: string; readonly selected: ReadonlySet<string> }[]
  >();
  const mutablePredicates = new Map<
    string,
    { readonly fieldId: string; readonly selected: ReadonlySet<string> }[]
  >();
  const slicerIds = new Set<string>();
  let targetCount = 0;
  for (const slicer of slicers) {
    if (!identifierPattern.test(slicer.id) || slicerIds.has(slicer.id)) {
      throw new TypeError(`Slicer ID ${slicer.id} is invalid or duplicated`);
    }
    slicerIds.add(slicer.id);
    if (!identifierPattern.test(slicer.fieldId)) {
      throw new TypeError(`Slicer ${slicer.id} field ID is invalid`);
    }
    const selection =
      slicer.stateScope === 'session' && sessionSelections[slicer.id] !== undefined
        ? sessionSelections[slicer.id]!
        : slicer.selection;
    if (selection.length > maximumSelections) {
      throw new RangeError(`Slicer ${slicer.id} selection value limit exceeded`);
    }
    const selected = new Set(selection.map(scalarKey));
    const targetIds = [...new Set(slicer.targets)];
    targetCount += targetIds.length;
    if (targetCount > maximumTargets) throw new RangeError('Slicer target limit exceeded');
    for (const targetId of targetIds) {
      if (!identifierPattern.test(targetId)) {
        throw new TypeError(`Slicer ${slicer.id} target ID is invalid`);
      }
      if (knownTargets !== undefined && !knownTargets.has(targetId)) {
        diagnostics.push(
          Object.freeze({
            code: 'SLICER_TARGET_MISSING',
            slicerId: slicer.id,
            targetId,
            message: `Slicer ${slicer.id} references missing target ${targetId}`,
          }),
        );
      }
      if (selected.size === 0) continue;
      const targetPredicates = mutablePredicates.get(targetId) ?? [];
      targetPredicates.push({ fieldId: slicer.fieldId, selected });
      mutablePredicates.set(targetId, targetPredicates);
    }
  }
  for (const [targetId, targetPredicates] of mutablePredicates) {
    predicates.set(targetId, Object.freeze([...targetPredicates]));
  }
  return Object.freeze({
    diagnostics: Object.freeze(diagnostics),
    matches(targetId: string, fields: Readonly<Record<string, PivotScalar>>): boolean {
      return (predicates.get(targetId) ?? []).every(({ fieldId, selected }) =>
        selected.has(scalarKey(fields[fieldId] ?? null)),
      );
    },
  });
}

/** Builds a deterministic, bounded unique-value index for Slicer creation UX. */
export function buildSlicerValueIndex(
  values: readonly PivotScalar[],
  options: SlicerValueIndexOptions = {},
): readonly SlicerValueIndexEntry[] {
  const maximumValues = limit(options.maximumValues, 10_000, 'Slicer value limit');
  const entries = new Map<string, { value: PivotScalar; count: number }>();
  for (const value of values) {
    const key = scalarKey(value);
    const existing = entries.get(key);
    if (existing === undefined) {
      if (entries.size >= maximumValues) throw new RangeError('Slicer value limit exceeded');
      entries.set(key, { value, count: 1 });
    } else {
      existing.count += 1;
    }
  }
  return Object.freeze(
    [...entries.values()]
      .sort((left, right) => compareScalar(left.value, right.value))
      .map((entry) => Object.freeze({ ...entry })),
  );
}
