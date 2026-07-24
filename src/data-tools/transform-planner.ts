import type {
  DocumentCommandEnvelope,
  DocumentController,
  DocumentControllerSnapshot,
  DocumentTransactionEnvelope,
  DocumentTransactionResult,
} from '../document-controller';
import type {
  CellInput,
  Diagnostic,
  DocumentCellAddress,
  DocumentCellRange,
  SparseCell,
} from '../document';
import type { SheetId } from '../core';
import { DataTransformError } from './errors';
import { indexSparseRange, yieldForCancellation } from './range-index';
import { createSafeRegexBudget, type SafeRegexBudget } from './safe-regex';

export { DataTransformError } from './errors';

/** Find-and-replace request over scalar text cells in one document range. */
export interface FindReplaceTransform {
  /** Transform discriminator. */
  readonly type: 'find-replace';
  /** Inclusive target range. */
  readonly range: DocumentCellRange;
  /** Literal text or regular-expression source to find. */
  readonly find: string;
  /** Replacement text. */
  readonly replacement: string;
  /** Explicit matching policy. */
  readonly match: 'literal' | 'regex';
}

/** Splits scalar text into adjacent columns without mutating formulas. */
export interface SplitTextTransform {
  /** Transform discriminator. */
  readonly type: 'split-text';
  /** Inclusive source range. */
  readonly range: DocumentCellRange;
  /** Non-empty literal delimiter. */
  readonly delimiter: string;
  /** Maximum output columns for any source cell. */
  readonly maximumColumns: number;
}

/** Removes duplicate worksheet rows using structural commands. */
export interface RemoveDuplicatesTransform {
  /** Transform discriminator. */
  readonly type: 'remove-duplicates';
  /** Inclusive row range inspected for duplicates. */
  readonly range: DocumentCellRange;
  /** Absolute worksheet columns forming the duplicate key. */
  readonly keyColumns: readonly number[];
  /** Which duplicate row is retained. */
  readonly keep: 'first' | 'last';
}

/** Generates a row-major numeric, ISO-date, or numeric-suffix text sequence. */
export interface FillSeriesTransform {
  /** Transform discriminator. */
  readonly type: 'fill-series';
  /** Inclusive row-major target range. */
  readonly range: DocumentCellRange;
  /** Explicit deterministic sequence category. */
  readonly series: 'number' | 'date' | 'text-suffix';
  /** One starting value or two values used to infer the step. */
  readonly seed: readonly string[];
}

/** Supported deterministic data-cleanup requests. */
export type DataTransform =
  | FindReplaceTransform
  | SplitTextTransform
  | RemoveDuplicatesTransform
  | FillSeriesTransform;

/** Immutable dry-run result bound to one document revision. */
export interface DataTransformPreview {
  /** Opaque identifier used to commit this plan. */
  readonly planId: string;
  /** Document revision used to build the plan. */
  readonly baseRevision: number;
  /** Exact logical range inspected by the planner. */
  readonly affectedRange: DocumentCellRange;
  /** Number of normalized commands in the complete plan. */
  readonly estimatedCellCount: number;
  /** Structured warnings that require user attention before commit. */
  readonly warnings: readonly Diagnostic[];
  /** Bounded preview of changed cells or deleted rows. */
  readonly sampleChanges: readonly {
    /** Zero-based source row. */
    readonly row: number;
    /** Zero-based source column. */
    readonly column: number;
    /** Cell text before the transform. */
    readonly before: string;
    /** Cell text after the transform. */
    readonly after: string;
  }[];
}

/** Commit result including planner-specific stale and missing failures. */
export type DataTransformCommitResult =
  | DocumentTransactionResult
  | {
      /** Indicates the plan cannot be committed. */
      readonly status: 'rejected';
      /** Stable planner rejection category. */
      readonly code: 'TRANSFORM_PLAN_MISSING' | 'TRANSFORM_PLAN_STALE';
    };

interface PendingPlan {
  readonly baseRevision: number;
  readonly commands: readonly DocumentCommandEnvelope[];
}

/** Cancellation options for an asynchronous transform preview. */
export interface DataTransformPreviewOptions {
  /** Cancels planning before a command plan is published. */
  readonly signal?: AbortSignal;
  /** Optional read-only template and calculated-value projections. */
  readonly context?: DataToolPreviewContext;
}

/** Optional host projections used only to add conservative preview/anomaly warnings. */
export interface DataToolPreviewContext {
  /** Template-owned regions that should warn when a transform intersects them. */
  readonly templateRegions?: readonly DocumentCellRange[];
  /** Cells whose calculated values are errors, when the host exposes that projection. */
  readonly errorCells?: readonly DocumentCellAddress[];
}

function inputText(input: CellInput): string {
  if (input.type === 'blank' || input.type === 'custom') return '';
  if (input.type === 'formula') return input.source;
  return String(input.value);
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function inRange(cell: Pick<SparseCell, 'row' | 'column'>, range: DocumentCellRange): boolean {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.column >= range.start.column &&
    cell.column <= range.end.column
  );
}

function snapshotRange(range: DocumentCellRange): DocumentCellRange {
  return Object.freeze({
    sheetId: range.sheetId,
    start: Object.freeze({ ...range.start }),
    end: Object.freeze({ ...range.end }),
  });
}

function snapshotTransform(transform: DataTransform): DataTransform {
  const range = snapshotRange(transform.range);
  if (transform.type === 'find-replace') {
    return Object.freeze({
      type: transform.type,
      range,
      find: String(transform.find),
      replacement: String(transform.replacement),
      match: transform.match,
    });
  }
  if (transform.type === 'split-text') {
    return Object.freeze({
      type: transform.type,
      range,
      delimiter: String(transform.delimiter),
      maximumColumns: transform.maximumColumns,
    });
  }
  if (transform.type === 'fill-series') {
    return Object.freeze({
      type: transform.type,
      range,
      series: transform.series,
      seed: Object.freeze(transform.seed.map(String)),
    });
  }
  return Object.freeze({
    type: transform.type,
    range,
    keyColumns: Object.freeze([...transform.keyColumns]),
    keep: transform.keep,
  });
}

function diagnostic(code: string, message: string, cell?: DocumentCellAddress): Diagnostic {
  return Object.freeze({
    code,
    severity: 'warning',
    domain: 'data',
    stage: 'plan',
    message,
    ...(cell === undefined
      ? {}
      : { location: Object.freeze({ cell: Object.freeze({ ...cell }) }) }),
  });
}

function rangeDiagnostic(code: string, message: string, range: DocumentCellRange): Diagnostic {
  return Object.freeze({
    code,
    severity: 'warning',
    domain: 'data',
    stage: 'plan',
    message,
    location: Object.freeze({ range: snapshotRange(range) }),
  });
}

function setCellCommand(
  sheet: SheetId,
  row: number,
  column: number,
  text: string,
  index: number,
): DocumentCommandEnvelope {
  return Object.freeze({
    schemaVersion: 1,
    id: `transform-cell-${index + 1}`,
    command: Object.freeze({
      type: 'set-cell-text',
      address: Object.freeze({ sheet, row, column }),
      text,
    }),
  });
}

function setCellInputCommand(
  sheet: SheetId,
  row: number,
  column: number,
  input: CellInput,
  index: number,
): DocumentCommandEnvelope {
  return Object.freeze({
    schemaVersion: 1,
    id: `transform-cell-${index + 1}`,
    command: Object.freeze({
      type: 'set-cell-input',
      address: Object.freeze({ sheet, row, column }),
      input: Object.freeze({ ...input }),
    }),
  });
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new DataTransformError('TRANSFORM_ABORTED', 'Data transform preview was aborted');
  }
}

function replacementPattern(
  transform: FindReplaceTransform,
  limits: {
    readonly maximumPatternLength: number;
    readonly maximumInputLength: number;
    readonly maximumSteps: number;
    readonly maximumMilliseconds: number;
  },
): SafeRegexBudget | undefined {
  if (transform.match === 'literal') return undefined;
  return createSafeRegexBudget(transform.find, limits);
}

function replaceText(
  value: string,
  transform: FindReplaceTransform,
  pattern: SafeRegexBudget | undefined,
): string {
  return pattern === undefined
    ? value.replaceAll(transform.find, transform.replacement)
    : pattern.replace(value, transform.replacement);
}

function duplicateKey(
  row: number,
  columns: readonly number[],
  cells: ReadonlyMap<string, SparseCell>,
): string {
  return JSON.stringify(
    columns.map((column) => cells.get(cellKey(row, column))?.cell.input ?? { type: 'blank' }),
  );
}

function parseIsoDate(value: string): number | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/u.exec(value);
  if (match === null) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
    ? timestamp
    : undefined;
}

function isoDate(timestamp: number): string {
  return new Date(timestamp).toISOString().slice(0, 10);
}

interface FillSeriesValue {
  readonly text: string;
  readonly input: CellInput;
}

function createFillSeries(transform: FillSeriesTransform): (index: number) => FillSeriesValue {
  if (transform.seed.length < 1 || transform.seed.length > 2) {
    throw new DataTransformError(
      'FILL_SERIES_INVALID',
      'Fill series requires one or two seed values',
    );
  }
  const [first, second] = transform.seed;
  if (first === undefined) {
    throw new DataTransformError('FILL_SERIES_INVALID', 'Fill series seed is missing');
  }
  if (transform.series === 'number') {
    const start = Number(first);
    const next = second === undefined ? start + 1 : Number(second);
    if (!Number.isFinite(start) || !Number.isFinite(next)) {
      throw new DataTransformError('FILL_SERIES_INVALID', 'Numeric fill seeds must be finite');
    }
    const step = next - start;
    return (index) => {
      const value = start + step * index;
      if (!Number.isFinite(value)) {
        throw new DataTransformError('FILL_SERIES_INVALID', 'Numeric fill result must be finite');
      }
      return Object.freeze({
        text: String(value),
        input: Object.freeze({ type: 'number', value }),
      });
    };
  }
  if (transform.series === 'date') {
    const start = parseIsoDate(first);
    const next =
      second === undefined
        ? start === undefined
          ? undefined
          : start + 86_400_000
        : parseIsoDate(second);
    if (start === undefined || next === undefined) {
      throw new DataTransformError(
        'FILL_SERIES_INVALID',
        'Date fill seeds must use valid YYYY-MM-DD values',
      );
    }
    const step = next - start;
    return (index) => {
      const text = isoDate(start + step * index);
      return Object.freeze({
        text,
        input: Object.freeze({ type: 'string', value: text }),
      });
    };
  }
  const firstMatch = /^(.*?)(-?\d+)$/u.exec(first);
  const secondMatch = second === undefined ? undefined : /^(.*?)(-?\d+)$/u.exec(second);
  if (
    firstMatch === null ||
    (second !== undefined && (secondMatch === null || secondMatch?.[1] !== firstMatch[1]))
  ) {
    throw new DataTransformError(
      'FILL_SERIES_INVALID',
      'Text-suffix seeds must share a prefix and end in an integer',
    );
  }
  const prefix = firstMatch[1] as string;
  const start = Number(firstMatch[2]);
  const next = secondMatch == null ? start + 1 : Number(secondMatch[2]);
  const width = Math.max(
    (firstMatch[2] as string).replace(/^-/, '').length,
    secondMatch?.[2]?.replace(/^-/, '').length ?? 0,
  );
  const step = next - start;
  return (index) => {
    const value = start + step * index;
    const digits = String(Math.abs(value)).padStart(width, '0');
    const text = `${prefix}${value < 0 ? '-' : ''}${digits}`;
    return Object.freeze({
      text,
      input: Object.freeze({ type: 'string', value: text }),
    });
  };
}

function rangesIntersect(left: DocumentCellRange, right: DocumentCellRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
}

/** Creates an isolated revision-bound preview and commit planner. */
export function createDataTransformPlanner(limits: {
  readonly maxCells: number;
  readonly maxCommands?: number;
  readonly maxSamples: number;
  /** Maximum regex source length; defaults to 1,000 UTF-16 code units. */
  readonly maxRegexPatternLength?: number;
  /** Maximum text length accepted by one regex match; defaults to 100,000. */
  readonly maxRegexInputLength?: number;
  /** Maximum conservative pattern-length × input-length budget per preview. */
  readonly maxRegexSteps?: number;
  /** Maximum cumulative synchronous regex execution time per preview. */
  readonly maxRegexMilliseconds?: number;
}): {
  /** Builds an immutable bounded preview against one document revision. */
  preview(
    snapshot: DocumentControllerSnapshot,
    transform: DataTransform,
    options?: DataTransformPreviewOptions,
  ): Promise<DataTransformPreview>;
  /** Atomically commits a previously previewed plan when its revision is still current. */
  commit(controller: DocumentController, planId: string): DataTransformCommitResult;
} {
  if (!Number.isSafeInteger(limits.maxCells) || limits.maxCells < 0) {
    throw new TypeError('maxCells must be a non-negative safe integer');
  }
  const maxCommands = limits.maxCommands ?? limits.maxCells;
  if (!Number.isSafeInteger(maxCommands) || maxCommands < 0) {
    throw new TypeError('maxCommands must be a non-negative safe integer');
  }
  if (!Number.isSafeInteger(limits.maxSamples) || limits.maxSamples < 0) {
    throw new TypeError('maxSamples must be a non-negative safe integer');
  }
  const regexLimits = Object.freeze({
    maximumPatternLength: limits.maxRegexPatternLength ?? 1_000,
    maximumInputLength: limits.maxRegexInputLength ?? 100_000,
    maximumSteps: limits.maxRegexSteps ?? 10_000_000,
    maximumMilliseconds: limits.maxRegexMilliseconds ?? 50,
  });
  if (
    !Number.isSafeInteger(regexLimits.maximumPatternLength) ||
    regexLimits.maximumPatternLength < 0 ||
    !Number.isSafeInteger(regexLimits.maximumInputLength) ||
    regexLimits.maximumInputLength < 0 ||
    !Number.isSafeInteger(regexLimits.maximumSteps) ||
    regexLimits.maximumSteps < 1 ||
    !Number.isFinite(regexLimits.maximumMilliseconds) ||
    regexLimits.maximumMilliseconds <= 0
  ) {
    throw new TypeError('Regex budgets must be finite non-negative limits');
  }
  const plans = new Map<string, PendingPlan>();
  let sequence = 0;
  return {
    async preview(snapshot, transformInput, options = {}) {
      throwIfAborted(options.signal);
      const transform = snapshotTransform(transformInput);
      const rowCount = transform.range.end.row - transform.range.start.row + 1;
      const columnCount = transform.range.end.column - transform.range.start.column + 1;
      if (rowCount < 1 || columnCount < 1 || rowCount > Math.floor(limits.maxCells / columnCount)) {
        throw new DataTransformError(
          'TRANSFORM_TOO_LARGE',
          'Data transform exceeds the configured cell limit',
        );
      }
      const sheet = snapshot.document.workbook.sheets.find(
        ({ id }) => id === transform.range.sheetId,
      );
      if (sheet === undefined) {
        throw new DataTransformError('TRANSFORM_TOO_LARGE', 'Data transform sheet does not exist');
      }
      const sheetId = transform.range.sheetId as string as SheetId;
      const indexedRange =
        transform.type === 'split-text'
          ? {
              sheetId: transform.range.sheetId,
              start: transform.range.start,
              end: {
                row: transform.range.end.row,
                column: transform.range.end.column + transform.maximumColumns - 1,
              },
            }
          : transform.range;
      if (!Number.isSafeInteger(indexedRange.end.column)) {
        throw new DataTransformError(
          'TEXT_SPLIT_OVERFLOW',
          'Text split output exceeds the supported column range',
        );
      }
      const cells = await indexSparseRange(sheet.cells, indexedRange, options.signal);
      const changes: { row: number; column: number; before: string; after: string }[] = [];
      const warnings: Diagnostic[] = [];
      const commands: DocumentCommandEnvelope[] = [];

      if (transform.type === 'find-replace') {
        const pattern = replacementPattern(transform, regexLimits);
        let inspected = 0;
        for (const entry of cells.values()) {
          inspected += 1;
          await yieldForCancellation(inspected, options.signal);
          const before = inputText(entry.cell.input);
          if (entry.cell.input.type === 'formula') {
            if (replaceText(before, transform, pattern) !== before) {
              warnings.push(
                diagnostic(
                  'FORMULA_TRANSFORM_SKIPPED',
                  'Formula source matched but was not rewritten by the data tool',
                  { sheetId: transform.range.sheetId, row: entry.row, column: entry.column },
                ),
              );
            }
            continue;
          }
          if (entry.cell.input.type !== 'string') continue;
          const after = replaceText(before, transform, pattern);
          if (before === after) continue;
          changes.push({ row: entry.row, column: entry.column, before, after });
          commands.push(setCellCommand(sheetId, entry.row, entry.column, after, commands.length));
          if (/^[=+@]/u.test(after)) {
            warnings.push(
              diagnostic('FORMULA_INJECTION_RISK', 'Replacement creates formula-like cell text', {
                sheetId: transform.range.sheetId,
                row: entry.row,
                column: entry.column,
              }),
            );
          }
        }
      } else if (transform.type === 'split-text') {
        if (
          transform.delimiter === '' ||
          !Number.isSafeInteger(transform.maximumColumns) ||
          transform.maximumColumns < 1
        ) {
          throw new DataTransformError(
            'TEXT_SPLIT_OVERFLOW',
            'Text split requires a delimiter and a positive column limit',
          );
        }
        let inspected = 0;
        for (const entry of cells.values()) {
          inspected += 1;
          await yieldForCancellation(inspected, options.signal);
          if (!inRange(entry, transform.range) || entry.cell.input.type !== 'string') continue;
          const parts = entry.cell.input.value.split(transform.delimiter);
          if (parts.length > transform.maximumColumns) {
            throw new DataTransformError(
              'TEXT_SPLIT_OVERFLOW',
              `Text split at row ${entry.row} exceeds ${transform.maximumColumns} columns`,
            );
          }
          for (const [offset, after] of parts.entries()) {
            const column = entry.column + offset;
            const existing = cells.get(cellKey(entry.row, column));
            const before = existing === undefined ? '' : inputText(existing.cell.input);
            if (before === after) continue;
            if (offset > 0 && before !== '') {
              warnings.push(
                diagnostic(
                  'NONEMPTY_TARGET_OVERWRITE',
                  'Text split overwrites a non-empty target cell',
                  { sheetId: transform.range.sheetId, row: entry.row, column },
                ),
              );
            }
            changes.push({ row: entry.row, column, before, after });
            commands.push(setCellCommand(sheetId, entry.row, column, after, commands.length));
          }
        }
      } else if (transform.type === 'remove-duplicates') {
        if (
          transform.keyColumns.length === 0 ||
          transform.keyColumns.some(
            (column) =>
              !Number.isSafeInteger(column) ||
              column < transform.range.start.column ||
              column > transform.range.end.column,
          )
        ) {
          throw new DataTransformError(
            'TRANSFORM_TOO_LARGE',
            'Duplicate-removal key columns must be inside the target range',
          );
        }
        const rows = Array.from(
          { length: rowCount },
          (_, index) => transform.range.start.row + index,
        );
        const scan = transform.keep === 'first' ? rows : [...rows].reverse();
        const retained = new Set<string>();
        const duplicates: number[] = [];
        let inspected = 0;
        for (const row of scan) {
          inspected += 1;
          await yieldForCancellation(inspected, options.signal);
          const key = duplicateKey(row, transform.keyColumns, cells);
          if (retained.has(key)) duplicates.push(row);
          else retained.add(key);
        }
        duplicates.sort((left, right) => right - left);
        for (const row of duplicates) {
          changes.push({
            row,
            column: transform.range.start.column,
            before: 'duplicate row',
            after: '',
          });
          commands.push(
            Object.freeze({
              schemaVersion: 1,
              id: `transform-row-${commands.length + 1}`,
              command: Object.freeze({ type: 'delete-row', sheet: sheetId, index: row, count: 1 }),
            }),
          );
        }
      } else {
        const valueAt = createFillSeries(transform);
        let index = 0;
        for (let row = transform.range.start.row; row <= transform.range.end.row; row += 1) {
          for (
            let column = transform.range.start.column;
            column <= transform.range.end.column;
            column += 1
          ) {
            await yieldForCancellation(index, options.signal);
            const value = valueAt(index);
            const after = value.text;
            index += 1;
            const existing = cells.get(cellKey(row, column));
            const before = existing === undefined ? '' : inputText(existing.cell.input);
            if (before === after) continue;
            if (
              existing !== undefined &&
              existing.cell.input.type !== 'blank' &&
              existing.cell.input.type !== (transform.series === 'number' ? 'number' : 'string')
            ) {
              warnings.push(
                diagnostic(
                  'FILL_SERIES_TYPE_OVERWRITE',
                  'Fill series overwrites a cell with a different input type',
                  { sheetId: transform.range.sheetId, row, column },
                ),
              );
            }
            changes.push({ row, column, before, after });
            commands.push(setCellInputCommand(sheetId, row, column, value.input, commands.length));
            if (/^[=+\-@]/u.test(after)) {
              warnings.push(
                diagnostic('FORMULA_INJECTION_RISK', 'Fill creates formula-like cell text', {
                  sheetId: transform.range.sheetId,
                  row,
                  column,
                }),
              );
            }
          }
        }
      }

      if (commands.length > maxCommands) {
        throw new DataTransformError(
          'TRANSFORM_TOO_LARGE',
          'Data transform exceeds the configured command limit',
        );
      }
      for (const region of options.context?.templateRegions ?? []) {
        if (!rangesIntersect(transform.range, region) || changes.length === 0) continue;
        warnings.push(
          rangeDiagnostic(
            'TEMPLATE_REGION_CONFLICT',
            'Data transform intersects a template-owned region',
            region,
          ),
        );
      }
      throwIfAborted(options.signal);
      const planId = `data-transform-${snapshot.revision}-${sequence}`;
      sequence += 1;
      plans.set(
        planId,
        Object.freeze({
          baseRevision: snapshot.revision,
          commands: Object.freeze([...commands]),
        }),
      );
      return Object.freeze({
        planId,
        baseRevision: snapshot.revision,
        affectedRange: snapshotRange(transform.range),
        estimatedCellCount: commands.length,
        warnings: Object.freeze([...warnings]),
        sampleChanges: Object.freeze(
          changes.slice(0, limits.maxSamples).map((change) => Object.freeze({ ...change })),
        ),
      });
    },
    commit(controller, planId) {
      const plan = plans.get(planId);
      if (plan === undefined) return { status: 'rejected', code: 'TRANSFORM_PLAN_MISSING' };
      plans.delete(planId);
      if (controller.getSnapshot().revision !== plan.baseRevision) {
        return { status: 'rejected', code: 'TRANSFORM_PLAN_STALE' };
      }
      const transaction: DocumentTransactionEnvelope = {
        schemaVersion: 1,
        id: planId,
        baseRevision: plan.baseRevision,
        commands: plan.commands,
      };
      return controller.transact(transaction);
    },
  };
}
