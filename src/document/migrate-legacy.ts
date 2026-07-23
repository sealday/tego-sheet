import type { JsonValue } from '../core/types/json';
import type { Diagnostic, DocumentDiagnosticCode } from './diagnostics';
import type {
  CellInput,
  SheetInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from './model/document';
import type { CellInputRecord } from './model/sparse-cells';
import { parseSpreadsheetDocument } from './parse-document';

/** Stable migration-specific diagnostic codes. */
export type LegacyMigrationDiagnosticCode =
  | 'LEGACY_FIELD_DROPPED'
  | 'LEGACY_FIELD_DEGRADED'
  | 'LEGACY_VALUE_INVALID'
  | 'LEGACY_REFERENCE_INVALID';

/** Diagnostic returned while migrating a legacy workbook. */
export interface MigrationDiagnostic extends Diagnostic {
  /** Stable migration or downstream schema validation code. */
  readonly code: LegacyMigrationDiagnosticCode | DocumentDiagnosticCode;
}

/** Caller-controlled factories for persistent IDs created by migration. */
export interface LegacyMigrationIdFactory {
  /** Creates the document ID exactly once. */
  readonly documentId: () => string;
  /** Creates a sheet ID in legacy sheet order. */
  readonly sheetId: (index: number, name: string) => string;
}

/** Options controlling legacy workbook migration. */
export interface LegacyMigrationOptions {
  /** Provide both factories for deterministic output. */
  readonly ids?: LegacyMigrationIdFactory;
  /** Date system assigned to legacy numeric date serials. */
  readonly dateSystem?: 'excel-1900' | 'excel-1904';
  /** Optional locale hint for migrated content. */
  readonly localeHint?: string;
}

/** Atomic result of legacy workbook migration. */
export type MigrationResult =
  | {
      /** Migration and schema validation succeeded. */
      readonly ok: true;
      /** Deeply frozen schema 2 document. */
      readonly document: SpreadsheetDocument;
      /** Non-fatal migration diagnostics. */
      readonly diagnostics: readonly MigrationDiagnostic[];
    }
  | {
      /** Migration failed atomically. */
      readonly ok: false;
      /** Errors with no partial document. */
      readonly diagnostics: readonly MigrationDiagnostic[];
    };

type RecordValue = Record<string, unknown>;
type MutableCell = CellInputRecord;
type Range = {
  start: { row: number; column: number };
  end: { row: number; column: number };
};
interface Context {
  diagnostics: MigrationDiagnostic[];
  styles: { id: string; value: JsonValue }[];
  validations: { id: string; value: JsonValue }[];
  styleKeys: Map<string, string>;
  validationKeys: Map<string, string>;
}

const SHEET = new Set([
  'name',
  'freeze',
  'styles',
  'merges',
  'rows',
  'cols',
  'validations',
  'autofilter',
]);
const ROW = new Set(['height', 'hide', 'style', 'cells']);
const COLUMN = new Set(['width', 'hide', 'style']);
const CELL = new Set(['text', 'type', 'value', 'style', 'merge', 'editable', 'printable']);
const STYLE = new Set([
  'format',
  'bgcolor',
  'align',
  'valign',
  'textwrap',
  'strike',
  'underline',
  'color',
  'font',
  'border',
]);
const FONT = new Set(['name', 'size', 'bold', 'italic']);
const BORDER = new Set(['top', 'right', 'bottom', 'left']);
const VALIDATION = new Set(['refs', 'mode', 'type', 'required', 'operator', 'value']);

function diag(
  code: LegacyMigrationDiagnosticCode,
  severity: 'warning' | 'error',
  path: string,
  message: string,
): MigrationDiagnostic {
  return {
    code,
    severity,
    domain: 'document',
    stage: 'migrate',
    message,
    details: { path },
  };
}

function fail(
  context: Context,
  code: 'LEGACY_VALUE_INVALID' | 'LEGACY_REFERENCE_INVALID',
  path: string,
  message: string,
): void {
  context.diagnostics.push(diag(code, 'error', path, message));
}

function warn(
  context: Context,
  code: 'LEGACY_FIELD_DROPPED' | 'LEGACY_FIELD_DEGRADED',
  path: string,
  message: string,
): void {
  context.diagnostics.push(diag(code, 'warning', path, message));
}

function record(value: unknown, path: string, context: Context): RecordValue | undefined {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be a plain object`);
    return undefined;
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be a plain object`);
    return undefined;
  }
  return value as RecordValue;
}

function unknown(
  value: RecordValue,
  allowed: ReadonlySet<string>,
  path: string,
  context: Context,
): void {
  for (const key of Object.keys(value).sort()) {
    if (!allowed.has(key)) {
      warn(context, 'LEGACY_FIELD_DROPPED', `${path}.${key}`, `${path}.${key} was dropped`);
    }
  }
}

function json(
  value: unknown,
  path: string,
  context: Context,
  active = new WeakSet<object>(),
): JsonValue | undefined {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean' ||
    (typeof value === 'number' && Number.isFinite(value))
  )
    return value;
  if (typeof value !== 'object' || value instanceof Date || active.has(value)) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be JSON-compatible`);
    return undefined;
  }
  active.add(value);
  try {
    if (Array.isArray(value)) {
      const output: JsonValue[] = [];
      for (let index = 0; index < value.length; index += 1) {
        const item = json(value[index], `${path}[${index}]`, context, active);
        if (item === undefined) return undefined;
        output.push(item);
      }
      return output;
    }
    const source = record(value, path, context);
    if (source === undefined) return undefined;
    const output = Object.create(null) as Record<string, JsonValue>;
    for (const key of Object.keys(source).sort()) {
      const item = json(source[key], `${path}.${key}`, context, active);
      if (item === undefined) return undefined;
      output[key] = item;
    }
    return output;
  } finally {
    active.delete(value);
  }
}

function indexes(
  value: RecordValue,
  reserved: ReadonlySet<string>,
  path: string,
  context: Context,
): readonly { key: string; index: number }[] {
  const output: { key: string; index: number }[] = [];
  for (const key of Object.keys(value)) {
    if (reserved.has(key)) continue;
    if (!/^(?:0|[1-9]\d*)$/.test(key) || !Number.isSafeInteger(Number(key))) {
      fail(context, 'LEGACY_VALUE_INVALID', `${path}.${key}`, `${path}.${key} is an invalid index`);
    } else output.push({ key, index: Number(key) });
  }
  return output.sort((left, right) => left.index - right.index);
}

function point(source: unknown): { row: number; column: number } | undefined {
  if (typeof source !== 'string') return undefined;
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)$/.exec(source);
  if (match === null) return undefined;
  let column = 0;
  for (const character of match[1]!.toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
    if (!Number.isSafeInteger(column)) return undefined;
  }
  const row = Number(match[2]) - 1;
  return Number.isSafeInteger(row) ? { row, column: column - 1 } : undefined;
}

function range(source: unknown): Range | undefined {
  if (typeof source !== 'string') return undefined;
  const parts = source.split(':');
  if (parts.length > 2) return undefined;
  const left = point(parts[0]);
  const right = point(parts[1] ?? parts[0]);
  if (left === undefined || right === undefined) return undefined;
  return {
    start: { row: Math.min(left.row, right.row), column: Math.min(left.column, right.column) },
    end: { row: Math.max(left.row, right.row), column: Math.max(left.column, right.column) },
  };
}

function styleIds(value: unknown, path: string, context: Context): readonly (string | undefined)[] {
  if (value === undefined) return [];
  if (!Array.isArray(value)) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be an array`);
    return [];
  }
  return value.map((entry, index) => {
    const entryPath = `${path}[${index}]`;
    const source = record(entry, entryPath, context);
    if (source === undefined) return undefined;
    unknown(source, STYLE, entryPath, context);
    const filtered = Object.fromEntries(
      [...STYLE].filter((key) => source[key] !== undefined).map((key) => [key, source[key]]),
    );
    for (const [field, allowed] of [
      ['font', FONT],
      ['border', BORDER],
    ] as const) {
      if (filtered[field] === undefined) continue;
      const nested = record(filtered[field], `${entryPath}.${field}`, context);
      if (nested === undefined) return undefined;
      unknown(nested, allowed, `${entryPath}.${field}`, context);
      filtered[field] = Object.fromEntries(
        [...allowed].filter((key) => nested[key] !== undefined).map((key) => [key, nested[key]]),
      );
    }
    const normalized = json(filtered, entryPath, context);
    if (normalized === undefined) return undefined;
    const key = JSON.stringify(normalized);
    const existing = context.styleKeys.get(key);
    if (existing !== undefined) return existing;
    const id = `legacy-style-${context.styles.length}`;
    context.styleKeys.set(key, id);
    context.styles.push({ id, value: normalized });
    return id;
  });
}

function inputAt(value: RecordValue, path: string, context: Context): CellInput | undefined {
  if (value.value instanceof Date) {
    fail(context, 'LEGACY_VALUE_INVALID', `${path}.value`, `${path}.value cannot persist a Date`);
    return undefined;
  }
  if (Object.hasOwn(value, 'text')) {
    if (typeof value.text !== 'string') {
      fail(context, 'LEGACY_VALUE_INVALID', `${path}.text`, `${path}.text must be a string`);
      return undefined;
    }
    return value.text.startsWith('=')
      ? { type: 'formula', source: value.text }
      : { type: 'string', value: value.text };
  }
  if (value.type === 'number') {
    if (typeof value.value !== 'number' || !Number.isFinite(value.value)) {
      fail(context, 'LEGACY_VALUE_INVALID', `${path}.value`, `${path}.value must be finite`);
      return undefined;
    }
    return { type: 'number', value: value.value };
  }
  if (value.type === 'boolean') {
    if (typeof value.value !== 'boolean') {
      fail(context, 'LEGACY_VALUE_INVALID', `${path}.value`, `${path}.value must be boolean`);
      return undefined;
    }
    return { type: 'boolean', value: value.value };
  }
  return { type: 'blank' };
}

function mergeList(value: unknown, path: string, context: Context): Range[] {
  if (value === undefined) return [];
  if (!Array.isArray(value)) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be an array`);
    return [];
  }
  return value.flatMap((entry, index) => {
    const parsed = range(entry);
    if (parsed !== undefined) return [parsed];
    fail(context, 'LEGACY_REFERENCE_INVALID', `${path}[${index}]`, `${path}[${index}] is invalid`);
    return [];
  });
}

function cells(
  value: unknown,
  ids: readonly (string | undefined)[],
  path: string,
  context: Context,
  merges: Range[],
): Map<string, { row: number; column: number; cell: MutableCell }> {
  const output = new Map<string, { row: number; column: number; cell: MutableCell }>();
  if (value === undefined) return output;
  const rows = record(value, path, context);
  if (rows === undefined) return output;
  if (rows.len !== undefined)
    warn(context, 'LEGACY_FIELD_DEGRADED', `${path}.len`, `${path}.len was not persisted`);
  for (const rowEntry of indexes(rows, new Set(['len']), path, context)) {
    const rowPath = `${path}.${rowEntry.key}`;
    const row = record(rows[rowEntry.key], rowPath, context);
    if (row === undefined) continue;
    unknown(row, ROW, rowPath, context);
    for (const field of ['height', 'hide', 'style'] as const)
      if (row[field] !== undefined)
        warn(
          context,
          'LEGACY_FIELD_DEGRADED',
          `${rowPath}.${field}`,
          `${rowPath}.${field} was not persisted`,
        );
    if (row.cells === undefined) continue;
    const cellPath = `${rowPath}.cells`;
    const source = record(row.cells, cellPath, context);
    if (source === undefined) continue;
    for (const columnEntry of indexes(source, new Set(), cellPath, context)) {
      const pathAt = `${cellPath}.${columnEntry.key}`;
      const legacy = record(source[columnEntry.key], pathAt, context);
      if (legacy === undefined) continue;
      unknown(legacy, CELL, pathAt, context);
      const input = inputAt(legacy, pathAt, context);
      if (input === undefined) continue;
      const cell: MutableCell = { input };
      if (legacy.style !== undefined) {
        const index = legacy.style;
        if (
          !Number.isSafeInteger(index) ||
          (index as number) < 0 ||
          ids[index as number] === undefined
        )
          fail(
            context,
            'LEGACY_REFERENCE_INVALID',
            `${pathAt}.style`,
            `${pathAt}.style is invalid`,
          );
        else cell.styleId = ids[index as number];
      }
      for (const field of ['editable', 'printable'] as const)
        if (legacy[field] !== undefined)
          warn(
            context,
            'LEGACY_FIELD_DEGRADED',
            `${pathAt}.${field}`,
            `${pathAt}.${field} was not persisted`,
          );
      if (legacy.merge !== undefined) {
        const spans = legacy.merge;
        if (
          !Array.isArray(spans) ||
          spans.length !== 2 ||
          !spans.every((span) => Number.isSafeInteger(span) && span >= 0)
        )
          fail(context, 'LEGACY_VALUE_INVALID', `${pathAt}.merge`, `${pathAt}.merge is invalid`);
        else
          merges.push({
            start: { row: rowEntry.index, column: columnEntry.index },
            end: {
              row: rowEntry.index + spans[0]!,
              column: columnEntry.index + spans[1]!,
            },
          });
      }
      output.set(`${rowEntry.index}:${columnEntry.index}`, {
        row: rowEntry.index,
        column: columnEntry.index,
        cell,
      });
    }
  }
  return output;
}

function columns(value: unknown, path: string, context: Context): void {
  if (value === undefined) return;
  const source = record(value, path, context);
  if (source === undefined) return;
  if (source.len !== undefined)
    warn(context, 'LEGACY_FIELD_DEGRADED', `${path}.len`, `${path}.len was not persisted`);
  for (const entry of indexes(source, new Set(['len']), path, context)) {
    const currentPath = `${path}.${entry.key}`;
    const column = record(source[entry.key], currentPath, context);
    if (column === undefined) continue;
    unknown(column, COLUMN, currentPath, context);
    for (const field of COLUMN)
      if (column[field] !== undefined)
        warn(
          context,
          'LEGACY_FIELD_DEGRADED',
          `${currentPath}.${field}`,
          `${currentPath}.${field} was not persisted`,
        );
  }
}

function validations(
  value: unknown,
  path: string,
  context: Context,
  target: Map<string, { row: number; column: number; cell: MutableCell }>,
): void {
  if (value === undefined) return;
  if (!Array.isArray(value)) {
    fail(context, 'LEGACY_VALUE_INVALID', path, `${path} must be an array`);
    return;
  }
  value.forEach((entry, index) => {
    const entryPath = `${path}[${index}]`;
    const source = record(entry, entryPath, context);
    if (source === undefined) return;
    unknown(source, VALIDATION, entryPath, context);
    if (!Array.isArray(source.refs) || !source.refs.every((item) => typeof item === 'string')) {
      fail(
        context,
        'LEGACY_REFERENCE_INVALID',
        `${entryPath}.refs`,
        `${entryPath}.refs is invalid`,
      );
      return;
    }
    const fields = Object.fromEntries(
      [...VALIDATION]
        .filter((key) => key !== 'refs' && source[key] !== undefined)
        .map((key) => [key, source[key]]),
    );
    const normalized = json(fields, entryPath, context);
    if (normalized === undefined) return;
    const canonical = JSON.stringify(normalized);
    let id = context.validationKeys.get(canonical);
    if (id === undefined) {
      id = `legacy-validation-${context.validations.length}`;
      context.validationKeys.set(canonical, id);
      context.validations.push({ id, value: normalized });
    }
    for (let refIndex = 0; refIndex < source.refs.length; refIndex += 1) {
      const parsed = range(source.refs[refIndex]);
      if (parsed === undefined) {
        fail(
          context,
          'LEGACY_REFERENCE_INVALID',
          `${entryPath}.refs[${refIndex}]`,
          `${entryPath}.refs[${refIndex}] is invalid`,
        );
        continue;
      }
      const count =
        (parsed.end.row - parsed.start.row + 1) * (parsed.end.column - parsed.start.column + 1);
      if (!Number.isSafeInteger(count) || count > 1_000_000) {
        fail(
          context,
          'LEGACY_REFERENCE_INVALID',
          `${entryPath}.refs[${refIndex}]`,
          `${entryPath}.refs[${refIndex}] is too large`,
        );
        continue;
      }
      for (let row = parsed.start.row; row <= parsed.end.row; row += 1)
        for (let column = parsed.start.column; column <= parsed.end.column; column += 1) {
          const key = `${row}:${column}`;
          const cell = target.get(key) ?? {
            row,
            column,
            cell: { input: { type: 'blank' } as const },
          };
          if (cell.cell.validationId === undefined) cell.cell.validationId = id;
          else if (cell.cell.validationId !== id)
            warn(
              context,
              'LEGACY_FIELD_DEGRADED',
              `${entryPath}.refs[${refIndex}]`,
              'overlapping validation preserved the first rule',
            );
          target.set(key, cell);
        }
    }
  });
}

function migrateSheet(
  value: unknown,
  index: number,
  ids: LegacyMigrationIdFactory,
  context: Context,
): SheetInput | undefined {
  const path = `$[${index}]`;
  const source = record(value, path, context);
  if (source === undefined) return undefined;
  unknown(source, SHEET, path, context);
  if (source.name !== undefined && typeof source.name !== 'string')
    fail(context, 'LEGACY_VALUE_INVALID', `${path}.name`, `${path}.name must be a string`);
  const name = typeof source.name === 'string' ? source.name : `sheet${index + 1}`;
  if (source.freeze !== undefined) {
    if (point(source.freeze) === undefined)
      fail(context, 'LEGACY_REFERENCE_INVALID', `${path}.freeze`, `${path}.freeze is invalid`);
    else
      warn(context, 'LEGACY_FIELD_DEGRADED', `${path}.freeze`, `${path}.freeze was not persisted`);
  }
  if (source.autofilter !== undefined) {
    if (record(source.autofilter, `${path}.autofilter`, context) !== undefined)
      warn(
        context,
        'LEGACY_FIELD_DEGRADED',
        `${path}.autofilter`,
        `${path}.autofilter was not persisted`,
      );
  }
  const localStyles = styleIds(source.styles, `${path}.styles`, context);
  const merges = mergeList(source.merges, `${path}.merges`, context);
  const sheetCells = cells(source.rows, localStyles, `${path}.rows`, context, merges);
  columns(source.cols, `${path}.cols`, context);
  validations(source.validations, `${path}.validations`, context, sheetCells);
  return {
    id: ids.sheetId(index, name),
    name,
    cells: [...sheetCells.values()].sort(
      (left, right) => left.row - right.row || left.column - right.column,
    ),
    merges,
  };
}

/**
 * Purely migrates a legacy single-sheet or ordered multi-sheet workbook.
 *
 * @param input - Legacy single-sheet object or ordered sheet array.
 * @param options - Optional deterministic IDs and workbook settings.
 * @returns An atomic schema 2 result with structured diagnostics.
 */
export function migrateLegacyWorkbook(
  input: unknown,
  options: LegacyMigrationOptions = {},
): MigrationResult {
  const context: Context = {
    diagnostics: [],
    styles: [],
    validations: [],
    styleKeys: new Map(),
    validationKeys: new Map(),
  };
  const ids = options.ids ?? {
    documentId: () => globalThis.crypto.randomUUID(),
    sheetId: () => globalThis.crypto.randomUUID(),
  };
  const sheets = (Array.isArray(input) ? input : [input])
    .map((sheet, index) => migrateSheet(sheet, index, ids, context))
    .filter((sheet): sheet is SheetInput => sheet !== undefined);
  if (context.diagnostics.some(({ severity }) => severity === 'error'))
    return { ok: false, diagnostics: Object.freeze(context.diagnostics) };
  const candidate: SpreadsheetDocumentInput = {
    schemaVersion: 2,
    id: ids.documentId(),
    workbook: {
      sheets,
      styles: context.styles,
      validations: context.validations,
      settings: {
        dateSystem: options.dateSystem ?? 'excel-1900',
        ...(options.localeHint === undefined ? {} : { localeHint: options.localeHint }),
      },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  };
  const parsed = parseSpreadsheetDocument(candidate);
  if (!parsed.ok)
    return {
      ok: false,
      diagnostics: Object.freeze(parsed.diagnostics.map((entry) => entry as MigrationDiagnostic)),
    };
  return {
    ok: true,
    document: parsed.document,
    diagnostics: Object.freeze(context.diagnostics),
  };
}
