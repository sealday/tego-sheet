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

/** Supported deterministic data-cleanup requests. */
export type DataTransform = FindReplaceTransform | SplitTextTransform | RemoveDuplicatesTransform;

/** Stable planning failure. */
export class DataTransformError extends Error {
  /** Creates a machine-readable planning failure. */
  constructor(
    readonly code:
      | 'TRANSFORM_ABORTED'
      | 'TRANSFORM_TOO_LARGE'
      | 'REPLACE_PATTERN_INVALID'
      | 'TEXT_SPLIT_OVERFLOW',
    message: string,
  ) {
    super(message);
    this.name = 'DataTransformError';
  }
}

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

interface PreviewOptions {
  readonly signal?: AbortSignal;
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

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new DataTransformError('TRANSFORM_ABORTED', 'Data transform preview was aborted');
  }
}

function replacementPattern(transform: FindReplaceTransform): RegExp | undefined {
  if (transform.match === 'literal') return undefined;
  if (transform.find.length > 1_000) {
    throw new DataTransformError('REPLACE_PATTERN_INVALID', 'Replacement pattern is too long');
  }
  try {
    return new RegExp(transform.find, 'gu');
  } catch (cause) {
    throw new DataTransformError(
      'REPLACE_PATTERN_INVALID',
      cause instanceof Error ? cause.message : 'Replacement pattern is invalid',
    );
  }
}

function replaceText(
  value: string,
  transform: FindReplaceTransform,
  pattern: RegExp | undefined,
): string {
  return pattern === undefined
    ? value.replaceAll(transform.find, transform.replacement)
    : value.replace(pattern, transform.replacement);
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

/** Creates an isolated revision-bound preview and commit planner. */
export function createDataTransformPlanner(limits: {
  readonly maxCells: number;
  readonly maxCommands?: number;
  readonly maxSamples: number;
}): {
  /** Builds an immutable bounded preview against one document revision. */
  preview(
    snapshot: DocumentControllerSnapshot,
    transform: DataTransform,
    options?: PreviewOptions,
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
      const cells = new Map(sheet.cells.map((entry) => [cellKey(entry.row, entry.column), entry]));
      const changes: { row: number; column: number; before: string; after: string }[] = [];
      const warnings: Diagnostic[] = [];
      const commands: DocumentCommandEnvelope[] = [];

      if (transform.type === 'find-replace') {
        const pattern = replacementPattern(transform);
        for (const entry of sheet.cells) {
          throwIfAborted(options.signal);
          if (!inRange(entry, transform.range)) continue;
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
        for (const entry of sheet.cells) {
          throwIfAborted(options.signal);
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
      } else {
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
        for (const row of scan) {
          throwIfAborted(options.signal);
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
      }

      if (commands.length > maxCommands) {
        throw new DataTransformError(
          'TRANSFORM_TOO_LARGE',
          'Data transform exceeds the configured command limit',
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
