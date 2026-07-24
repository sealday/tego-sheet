import type { JsonValue } from '../core/types/json';
import type {
  DocumentDiagnostic,
  DocumentDiagnosticCode,
  DiagnosticDomain,
  DiagnosticStage,
} from './diagnostics';
import type {
  Cell,
  CellInput,
  CellPoint,
  ConditionalFormat,
  ConditionalStyle,
  ExtensionStore,
  FilterView,
  FilterViewPredicate,
  ObjectAnchor,
  ResourceMetadata,
  Sheet,
  SheetColumn,
  SheetFilter,
  SheetFilterItem,
  SheetGroup,
  SheetRange,
  SheetRow,
  SheetObject,
  StructuredTable,
  StructuredTableColumn,
  SpreadsheetDocument,
  StoredSpreadsheetTemplate,
} from './model/document';
import type { TemplateBinding, TemplatePrintProfile } from '../template/model';
import type { ChartDefinition, ChartSeriesDefinition, ChartType } from '../analysis/charts';
import type { SparklineDefinition, SparklineType } from '../analysis/sparklines';
import { BUILTIN_FORMULA_COMPATIBILITY, parseFormula, type FormulaAst } from '../formula';
import type {
  DocumentId,
  DocumentSheetId,
  GroupId,
  ObjectId,
  ResourceId,
  StyleId,
  TableColumnId,
  TableId,
  TemplateId,
  ValidationId,
} from './model/ids';
import { compareSparseCells } from './model/sparse-cells';
import type { SparseCell } from './model/sparse-cells';
import { hasActiveStructuredTableProjection } from './model/structured-table-projection';

/** Configurable safety limits enforced before deep document decoding. */
export interface DocumentLimits {
  /** Maximum number of sheets. */
  readonly maxSheets?: number;
  /** Maximum total number of sparse cells. */
  readonly maxCells?: number;
  /** Maximum total number of merge ranges. */
  readonly maxMerges?: number;
  /** Maximum total number of sparse row layout entries. */
  readonly maxRows?: number;
  /** Maximum total number of sparse column layout entries. */
  readonly maxColumns?: number;
  /** Maximum total number of saved view definitions. */
  readonly maxViews?: number;
  /** Maximum total number of floating objects. */
  readonly maxObjects?: number;
  /** Maximum total number of outline groups. */
  readonly maxGroups?: number;
  /** Maximum total number of structured tables. */
  readonly maxTables?: number;
  /** Maximum total number of structured table columns. */
  readonly maxTableColumns?: number;
  /** Maximum UTF-8 byte size of the input document. */
  readonly maxBytes?: number;
}

/** Options controlling Workbook 2.0 parsing. */
export interface DocumentParseOptions {
  /** Optional safety-limit overrides. */
  readonly limits?: DocumentLimits;
}

/** Atomic result of parsing and validating a Workbook 2.0 document. */
export type DocumentParseResult =
  | {
      /** Indicates that parsing succeeded. */
      readonly ok: true;
      /** Deeply frozen parsed document snapshot. */
      readonly document: SpreadsheetDocument;
      /** Successful parses contain no diagnostics. */
      readonly diagnostics: readonly [];
    }
  | {
      /** Indicates that parsing failed atomically. */
      readonly ok: false;
      /** Aggregated structured diagnostics; no partial document is exposed. */
      readonly diagnostics: readonly DocumentDiagnostic[];
    };

const DEFAULT_LIMITS = {
  maxSheets: 1_000,
  maxCells: 1_000_000,
  maxMerges: 2_000,
  maxRows: 1_000_000,
  maxColumns: 1_000_000,
  maxViews: 10_000,
  maxObjects: 100_000,
  maxGroups: 10_000,
  maxTables: 10_000,
  maxTableColumns: 100_000,
  maxBytes: 64 * 1024 * 1024,
} as const;
const NAMESPACE_PATTERN = /^[a-z0-9]+(?:[.-][a-z0-9]+)+$/i;
const nativeObjectConstructorSource = Function.prototype.toString.call(Object);

function isPlainObjectPrototype(prototype: object | null): boolean {
  if (prototype === null || prototype === Object.prototype) return true;
  if (Object.getPrototypeOf(prototype) !== null) return false;
  const constructor = Object.getOwnPropertyDescriptor(prototype, 'constructor')?.value;
  return (
    typeof constructor === 'function' &&
    Function.prototype.toString.call(constructor) === nativeObjectConstructorSource
  );
}

function deepFreeze<T>(value: T, seen = new WeakSet<object>()): T {
  if (value === null || typeof value !== 'object' || seen.has(value)) return value;
  seen.add(value);
  for (const item of Object.values(value as Record<string, unknown>)) {
    deepFreeze(item, seen);
  }
  return Object.freeze(value);
}

function canonicalJson(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(canonicalJson);
  if (value !== null && typeof value === 'object') {
    const source = value as { readonly [key: string]: JsonValue };
    const output = Object.create(null) as Record<string, JsonValue>;
    for (const key of Object.keys(source).sort()) output[key] = canonicalJson(source[key]!);
    return output;
  }
  return value;
}

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

interface ParseContext {
  readonly diagnostics: DocumentDiagnostic[];
  readonly activeJson: WeakSet<object>;
}

interface ResolvedDocumentLimits {
  readonly maxSheets: number;
  readonly maxCells: number;
  readonly maxMerges: number;
  readonly maxRows: number;
  readonly maxColumns: number;
  readonly maxViews: number;
  readonly maxObjects: number;
  readonly maxGroups: number;
  readonly maxTables: number;
  readonly maxTableColumns: number;
  readonly maxBytes: number;
}

const LIMIT_NAMES = [
  'maxSheets',
  'maxCells',
  'maxMerges',
  'maxRows',
  'maxColumns',
  'maxViews',
  'maxObjects',
  'maxGroups',
  'maxTables',
  'maxTableColumns',
  'maxBytes',
] as const;

function resolveLimits(
  options: DocumentParseOptions,
  context: ParseContext,
): ResolvedDocumentLimits | undefined {
  const limits = { ...DEFAULT_LIMITS };
  for (const name of LIMIT_NAMES) {
    const value = options.limits?.[name];
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value < 0) {
      addDiagnostic(
        context,
        'DOCUMENT_LIMIT_EXCEEDED',
        `$.limits.${name}`,
        `${name} must be a non-negative safe integer`,
      );
      return undefined;
    }
    Object.assign(limits, { [name]: value });
  }
  return limits;
}

class InputCaptureError extends Error {
  constructor(
    readonly code:
      | 'DOCUMENT_SCHEMA_INVALID'
      | 'DOCUMENT_LIMIT_EXCEEDED'
      | 'GROUP_LIMIT_EXCEEDED'
      | 'TABLE_LIMIT_EXCEEDED'
      | 'INVALID_EXTENSION_DATA',
    readonly path: string,
    message: string,
  ) {
    super(message);
  }
}

interface InputCaptureContext {
  readonly active: WeakSet<object>;
  readonly encoder: TextEncoder;
  readonly limits: ResolvedDocumentLimits;
  bytes: number;
  cells: number;
  merges: number;
  rows: number;
  columns: number;
  views: number;
  objects: number;
  groups: number;
  tables: number;
  tableColumns: number;
}

function captureInput(input: unknown, limits: ResolvedDocumentLimits): unknown {
  const context: InputCaptureContext = {
    active: new WeakSet(),
    encoder: new TextEncoder(),
    limits,
    bytes: 0,
    cells: 0,
    merges: 0,
    rows: 0,
    columns: 0,
    views: 0,
    objects: 0,
    groups: 0,
    tables: 0,
    tableColumns: 0,
  };

  const consume = (text: string): void => {
    context.bytes += context.encoder.encode(text).byteLength;
    if (context.bytes > context.limits.maxBytes) {
      throw new InputCaptureError(
        'DOCUMENT_LIMIT_EXCEEDED',
        '$',
        '$ exceeds its configured document limit',
      );
    }
  };
  const invalidJsonCode = (path: string): InputCaptureError['code'] =>
    path.startsWith('$.extensions.') ? 'INVALID_EXTENSION_DATA' : 'DOCUMENT_SCHEMA_INVALID';
  const capture = (value: unknown, path: string): unknown => {
    if (value === null || typeof value === 'string' || typeof value === 'boolean') {
      consume(JSON.stringify(value) ?? 'null');
      return value;
    }
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) {
        throw new InputCaptureError(
          invalidJsonCode(path),
          path,
          `${path} must contain a finite JSON number`,
        );
      }
      consume(JSON.stringify(value));
      return value;
    }
    if (typeof value !== 'object') {
      throw new InputCaptureError(invalidJsonCode(path), path, `${path} must contain a JSON value`);
    }
    if (context.active.has(value)) {
      throw new InputCaptureError(
        'DOCUMENT_SCHEMA_INVALID',
        path,
        `${path} must not contain circular references`,
      );
    }

    context.active.add(value);
    try {
      const descriptors = Object.getOwnPropertyDescriptors(value);

      if (Array.isArray(value)) {
        const length = descriptors.length?.value;
        if (typeof length !== 'number') {
          throw new InputCaptureError(
            'DOCUMENT_SCHEMA_INVALID',
            path,
            `${path} must have a valid array length`,
          );
        }
        if (path === '$.workbook.sheets' && length > context.limits.maxSheets) {
          throw new InputCaptureError(
            'DOCUMENT_LIMIT_EXCEEDED',
            path,
            `${path} exceeds its configured document limit`,
          );
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.cells$/.test(path)) {
          context.cells += length;
          if (context.cells > context.limits.maxCells) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured document limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.merges$/.test(path)) {
          context.merges += length;
          if (context.merges > context.limits.maxMerges) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured document limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.rows$/.test(path)) {
          context.rows += length;
          if (context.rows > context.limits.maxRows) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured row layout limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.columns$/.test(path)) {
          context.columns += length;
          if (context.columns > context.limits.maxColumns) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured column layout limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.filterViews$/.test(path)) {
          context.views += length;
          if (context.views > context.limits.maxViews) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured saved-view limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.objects$/.test(path)) {
          context.objects += length;
          if (context.objects > context.limits.maxObjects) {
            throw new InputCaptureError(
              'DOCUMENT_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured object limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.groups$/.test(path)) {
          context.groups += length;
          if (context.groups > context.limits.maxGroups) {
            throw new InputCaptureError(
              'GROUP_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured outline group limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.tables$/.test(path)) {
          context.tables += length;
          if (context.tables > context.limits.maxTables) {
            throw new InputCaptureError(
              'TABLE_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured structured table limit',
            );
          }
        }
        if (/^\$\.workbook\.sheets\[\d+\]\.tables\[\d+\]\.columns$/.test(path)) {
          context.tableColumns += length;
          if (context.tableColumns > context.limits.maxTableColumns) {
            throw new InputCaptureError(
              'TABLE_LIMIT_EXCEEDED',
              '$.workbook.sheets',
              '$.workbook.sheets exceeds its configured structured table column limit',
            );
          }
        }

        consume('[');
        const output = Array.from<unknown>({ length });
        for (let index = 0; index < length; index += 1) {
          if (index > 0) consume(',');
          const descriptor = descriptors[index];
          if (descriptor === undefined) {
            consume('null');
            continue;
          }
          if (!('value' in descriptor)) {
            throw new InputCaptureError(
              'DOCUMENT_SCHEMA_INVALID',
              `${path}[${index}]`,
              `${path}[${index}] must be a data property`,
            );
          }
          output[index] = capture(descriptor.value, `${path}[${index}]`);
        }
        consume(']');
        for (const [key, descriptor] of Object.entries(descriptors)) {
          if (descriptor.enumerable && !('value' in descriptor)) {
            throw new InputCaptureError(
              'DOCUMENT_SCHEMA_INVALID',
              `${path}.${key}`,
              `${path}.${key} must be a data property`,
            );
          }
        }
        return output;
      }

      const prototype = Object.getPrototypeOf(value);
      if (!isPlainObjectPrototype(prototype)) {
        throw new InputCaptureError(
          'DOCUMENT_SCHEMA_INVALID',
          path,
          `${path} must be a plain JSON object`,
        );
      }
      consume('{');
      let emitted = 0;
      const output = Object.create(null) as Record<string, unknown>;
      for (const [key, descriptor] of Object.entries(descriptors)) {
        if (!descriptor.enumerable) continue;
        if (!('value' in descriptor)) {
          throw new InputCaptureError(
            'DOCUMENT_SCHEMA_INVALID',
            `${path}.${key}`,
            `${path}.${key} must be a data property`,
          );
        }
        const item = descriptor.value;
        Object.defineProperty(output, key, {
          configurable: true,
          enumerable: true,
          writable: true,
          value: capture(item, `${path}.${key}`),
        });
        if (item === undefined || typeof item === 'function' || typeof item === 'symbol') continue;
        if (emitted > 0) consume(',');
        emitted += 1;
        consume(JSON.stringify(key));
        consume(':');
      }
      consume('}');
      return output;
    } finally {
      context.active.delete(value);
    }
  };

  return capture(input, '$');
}

function exceedsCollectionLimits(
  input: unknown,
  limits: ResolvedDocumentLimits,
): string | undefined {
  if (!isRecord(input) || !isRecord(input.workbook) || !Array.isArray(input.workbook.sheets)) {
    return undefined;
  }
  const sheets = input.workbook.sheets;
  if (sheets.length > limits.maxSheets) return '$.workbook.sheets';
  let cells = 0;
  let merges = 0;
  let rows = 0;
  let columns = 0;
  let views = 0;
  let objects = 0;
  for (const sheet of sheets) {
    if (!isRecord(sheet)) continue;
    if (Array.isArray(sheet.cells)) {
      cells += sheet.cells.length;
      if (cells > limits.maxCells) return '$.workbook.sheets';
    }
    if (Array.isArray(sheet.merges)) {
      merges += sheet.merges.length;
      if (merges > limits.maxMerges) return '$.workbook.sheets';
    }
    if (Array.isArray(sheet.rows)) {
      rows += sheet.rows.length;
      if (rows > limits.maxRows) return '$.workbook.sheets';
    }
    if (Array.isArray(sheet.columns)) {
      columns += sheet.columns.length;
      if (columns > limits.maxColumns) return '$.workbook.sheets';
    }
    if (Array.isArray(sheet.filterViews)) {
      views += sheet.filterViews.length;
      if (views > limits.maxViews) return '$.workbook.sheets';
    }
    if (Array.isArray(sheet.objects)) {
      objects += sheet.objects.length;
      if (objects > limits.maxObjects) return '$.workbook.sheets';
    }
  }
  return undefined;
}

const isRecord = (value: unknown): value is Record<string, unknown> => {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  return isPlainObjectPrototype(Object.getPrototypeOf(value));
};

function addDiagnostic(
  context: ParseContext,
  code: DocumentDiagnosticCode,
  path: string,
  message: string,
  domain: DiagnosticDomain = 'document',
  stage: DiagnosticStage = 'validate',
): void {
  context.diagnostics.push({
    code,
    severity: 'error',
    domain,
    stage,
    message,
    details: { path },
  });
}

function recordAt(
  value: unknown,
  path: string,
  context: ParseContext,
): Record<string, unknown> | undefined {
  if (!isRecord(value)) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a JSON object`,
      'document',
      'decode',
    );
    return undefined;
  }
  return value;
}

function arrayAt(value: unknown, path: string, context: ParseContext): readonly unknown[] {
  if (!Array.isArray(value)) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be an array`,
      'document',
      'decode',
    );
    return [];
  }
  return value;
}

function stringAt(value: unknown, path: string, context: ParseContext): string {
  if (typeof value !== 'string' || value.length === 0) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a non-empty string`,
      'document',
      'decode',
    );
    return '';
  }
  return value;
}

function displayStringAt(value: unknown, path: string, context: ParseContext): string {
  if (typeof value !== 'string') {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a string`,
      'document',
      'decode',
    );
    return '';
  }
  return value;
}

function finiteAt(value: unknown, path: string, context: ParseContext): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a finite number`,
      'document',
      'decode',
    );
    return 0;
  }
  return value;
}

function nonNegativeFiniteAt(value: unknown, path: string, context: ParseContext): number {
  const result = finiteAt(value, path, context);
  if (result < 0) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be non-negative`,
      'document',
      'decode',
    );
  }
  return result;
}

function indexAt(value: unknown, path: string, context: ParseContext): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a non-negative safe integer`,
      'document',
      'decode',
    );
    return 0;
  }
  return value;
}

function jsonAt(
  value: unknown,
  path: string,
  context: ParseContext,
  invalidCode: DocumentDiagnosticCode = 'DOCUMENT_SCHEMA_INVALID',
): JsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    if (Number.isFinite(value)) return value;
    addDiagnostic(
      context,
      invalidCode,
      path,
      `${path} must contain finite JSON numbers`,
      invalidCode === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
      'decode',
    );
    return null;
  }
  if (typeof value !== 'object') {
    addDiagnostic(
      context,
      invalidCode,
      path,
      `${path} must be JSON-compatible`,
      invalidCode === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
      'decode',
    );
    return null;
  }
  if (context.activeJson.has(value)) {
    addDiagnostic(
      context,
      invalidCode,
      path,
      `${path} must not contain circular references`,
      invalidCode === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
      'decode',
    );
    return null;
  }
  context.activeJson.add(value);
  try {
    if (Array.isArray(value)) {
      const output: JsonValue[] = [];
      for (let index = 0; index < value.length; index += 1) {
        if (!Object.prototype.hasOwnProperty.call(value, index)) {
          addDiagnostic(
            context,
            invalidCode,
            `${path}[${index}]`,
            `${path} must not contain sparse array holes`,
            invalidCode === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
            'decode',
          );
          output.push(null);
        } else {
          output.push(jsonAt(value[index], `${path}[${index}]`, context, invalidCode));
        }
      }
      return output;
    }
    if (!isRecord(value)) {
      addDiagnostic(
        context,
        invalidCode,
        path,
        `${path} must be a plain JSON object`,
        invalidCode === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
        'decode',
      );
      return null;
    }
    const record = value;
    const output = Object.create(null) as Record<string, JsonValue>;
    for (const key of Object.keys(record).sort()) {
      output[key] = jsonAt(record[key], `${path}.${key}`, context, invalidCode);
    }
    return output;
  } finally {
    context.activeJson.delete(value);
  }
}

function pointAt(value: unknown, path: string, context: ParseContext): CellPoint {
  const record = recordAt(value, path, context);
  return {
    row: indexAt(record?.row, `${path}.row`, context),
    column: indexAt(record?.column, `${path}.column`, context),
  };
}

function rangeAt(value: unknown, path: string, context: ParseContext): SheetRange {
  const record = recordAt(value, path, context);
  return {
    start: pointAt(record?.start, `${path}.start`, context),
    end: pointAt(record?.end, `${path}.end`, context),
  };
}

function isNormalized(range: SheetRange): boolean {
  return range.start.row <= range.end.row && range.start.column <= range.end.column;
}

function cellInputAt(value: unknown, path: string, context: ParseContext): CellInput {
  const record = recordAt(value, path, context) ?? {};
  const type = record.type;
  switch (type) {
    case 'blank':
      return { type };
    case 'string':
      if (typeof record.value !== 'string') {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${path}.value`,
          `${path}.value must be a string`,
          'document',
          'decode',
        );
      }
      return { type, value: typeof record.value === 'string' ? record.value : '' };
    case 'number':
      return { type, value: finiteAt(record.value, `${path}.value`, context) };
    case 'boolean':
      if (typeof record.value !== 'boolean') {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${path}.value`,
          `${path}.value must be a boolean`,
          'document',
          'decode',
        );
      }
      return { type, value: typeof record.value === 'boolean' ? record.value : false };
    case 'formula':
      return { type, source: stringAt(record.source, `${path}.source`, context) };
    case 'custom':
      return {
        type,
        cellType: stringAt(record.cellType, `${path}.cellType`, context),
        schemaVersion: indexAt(record.schemaVersion, `${path}.schemaVersion`, context),
        value: canonicalJson(jsonAt(record.value, `${path}.value`, context)),
      };
    default:
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${path}.type`,
        `${path}.type is not a supported cell input`,
        'document',
        'decode',
      );
      return { type: 'blank' };
  }
}

function cellAt(value: unknown, path: string, context: ParseContext): Cell {
  const record = recordAt(value, path, context);
  return {
    input: cellInputAt(record?.input, `${path}.input`, context),
    ...(record?.styleId === undefined
      ? {}
      : { styleId: stringAt(record.styleId, `${path}.styleId`, context) as StyleId }),
    ...(record?.validationId === undefined
      ? {}
      : {
          validationId: stringAt(
            record.validationId,
            `${path}.validationId`,
            context,
          ) as ValidationId,
        }),
    ...(record?.resourceId === undefined
      ? {}
      : {
          resourceId: stringAt(record.resourceId, `${path}.resourceId`, context) as ResourceId,
        }),
    ...(record?.templateId === undefined
      ? {}
      : {
          templateId: stringAt(record.templateId, `${path}.templateId`, context) as TemplateId,
        }),
    ...(record?.metadata === undefined
      ? {}
      : { metadata: canonicalJson(jsonAt(record.metadata, `${path}.metadata`, context)) }),
    ...(record?.editable === undefined
      ? {}
      : { editable: booleanAt(record.editable, `${path}.editable`, context) }),
    ...(record?.printable === undefined
      ? {}
      : { printable: booleanAt(record.printable, `${path}.printable`, context) }),
  };
}

function booleanAt(value: unknown, path: string, context: ParseContext): boolean {
  if (typeof value !== 'boolean') {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a boolean`,
      'document',
      'decode',
    );
    return false;
  }
  return value;
}

function layoutAt<T extends SheetRow | SheetColumn>(
  value: unknown,
  path: string,
  context: ParseContext,
  sizeField: 'height' | 'width',
): T[] {
  const seen = new Set<number>();
  return arrayAt(value ?? [], path, context).map((entry, index) => {
    const entryPath = `${path}[${index}]`;
    const item = recordAt(entry, entryPath, context);
    const layoutIndex = indexAt(item?.index, `${entryPath}.index`, context);
    if (seen.has(layoutIndex)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        entryPath,
        `${path} indexes must be unique`,
      );
    }
    seen.add(layoutIndex);
    return {
      index: layoutIndex,
      ...(item?.[sizeField] === undefined
        ? {}
        : {
            [sizeField]: nonNegativeFiniteAt(item[sizeField], `${entryPath}.${sizeField}`, context),
          }),
      ...(item?.hidden === undefined
        ? {}
        : { hidden: booleanAt(item.hidden, `${entryPath}.hidden`, context) }),
      ...(item?.styleId === undefined
        ? {}
        : { styleId: stringAt(item.styleId, `${entryPath}.styleId`, context) as StyleId }),
    } as T;
  });
}

const MAX_GROUP_LEVEL = 8;

function groupsAt(
  value: unknown,
  path: string,
  context: ParseContext,
  rowCount: number | undefined,
  columnCount: number | undefined,
): SheetGroup[] {
  const seenIds = new Set<string>();
  const groups = arrayAt(value ?? [], path, context).map((entry, index) => {
    const entryPath = `${path}[${index}]`;
    const item = recordAt(entry, entryPath, context);
    const id = stringAt(item?.id, `${entryPath}.id`, context);
    if (id.length === 0) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.id`,
        'Outline group id must not be empty',
      );
    }
    if (seenIds.has(id)) {
      addDiagnostic(context, 'DUPLICATE_ID', `${entryPath}.id`, `Duplicate stable ID ${id}`);
    }
    seenIds.add(id);
    const axis: SheetGroup['axis'] =
      item?.axis === 'row' || item?.axis === 'column' ? item.axis : 'row';
    if (item?.axis !== axis) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.axis`,
        'Outline group axis must be row or column',
      );
    }
    const start = indexAt(item?.start, `${entryPath}.start`, context);
    const end = indexAt(item?.end, `${entryPath}.end`, context);
    indexAt(item?.level, `${entryPath}.level`, context);
    const collapsed = booleanAt(item?.collapsed, `${entryPath}.collapsed`, context);
    const logicalCount = axis === 'row' ? rowCount : columnCount;
    if (start > end || (logicalCount !== undefined && end >= logicalCount)) {
      addDiagnostic(
        context,
        'GROUP_LIMIT_EXCEEDED',
        entryPath,
        'Outline group range must be non-empty and inside the logical worksheet size',
        'data',
      );
    }
    return { id: id as GroupId, axis, start, end, level: 1, collapsed };
  });

  const output: SheetGroup[] = [];
  for (const axis of ['row', 'column'] as const) {
    const axisGroups = groups
      .filter((group) => group.axis === axis)
      .sort(
        (left, right) =>
          left.start - right.start || right.end - left.end || compareCodeUnits(left.id, right.id),
      );
    const stack: SheetGroup[] = [];
    for (const group of axisGroups) {
      while (stack.length > 0 && group.start > stack.at(-1)!.end) stack.pop();
      let crossing = false;
      while (stack.length > 0 && group.end > stack.at(-1)!.end) {
        crossing = true;
        stack.pop();
      }
      if (crossing) {
        addDiagnostic(
          context,
          'GROUP_LIMIT_EXCEEDED',
          path,
          'Outline groups on one axis must be disjoint or properly nested',
          'data',
        );
      }
      const level = stack.length + 1;
      if (level > MAX_GROUP_LEVEL) {
        addDiagnostic(
          context,
          'GROUP_LIMIT_EXCEEDED',
          path,
          `Outline nesting cannot exceed ${MAX_GROUP_LEVEL} levels`,
          'data',
        );
      }
      const normalized = { ...group, level };
      output.push(normalized);
      stack.push(normalized);
    }
  }
  return output;
}

function filterAt(value: unknown, path: string, context: ParseContext): SheetFilter | undefined {
  if (value === undefined) return undefined;
  const source = recordAt(value, path, context);
  const filterItems = arrayAt(source?.filters, `${path}.filters`, context).map(
    (entry, index): SheetFilterItem => {
      const entryPath = `${path}.filters[${index}]`;
      const item = recordAt(entry, entryPath, context);
      const operator = item?.operator === 'all' || item?.operator === 'in' ? item.operator : 'all';
      if (item?.operator !== operator) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${entryPath}.operator`,
          `${entryPath}.operator must be all or in`,
        );
      }
      return {
        column: indexAt(item?.column, `${entryPath}.column`, context),
        operator,
        values: arrayAt(item?.values, `${entryPath}.values`, context).map((entry, valueIndex) =>
          displayStringAt(entry, `${entryPath}.values[${valueIndex}]`, context),
        ),
      };
    },
  );
  let sort: SheetFilter['sort'];
  if (source?.sort === null) sort = null;
  else if (source?.sort !== undefined) {
    const item = recordAt(source.sort, `${path}.sort`, context);
    const direction =
      item?.direction === 'asc' || item?.direction === 'desc' ? item.direction : 'asc';
    if (item?.direction !== direction) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${path}.sort.direction`,
        `${path}.sort.direction must be asc or desc`,
      );
    }
    sort = {
      column: indexAt(item?.column, `${path}.sort.column`, context),
      direction,
    };
  }
  const range =
    source?.range === undefined ? undefined : rangeAt(source.range, `${path}.range`, context);
  if (range !== undefined && !isNormalized(range)) {
    addDiagnostic(context, 'INVALID_RANGE', `${path}.range`, 'Filter range must be normalized');
  }
  return {
    ...(range === undefined ? {} : { range }),
    filters: filterItems,
    ...(sort === undefined ? {} : { sort }),
  };
}

const CONDITIONAL_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
]);
const CONDITIONAL_STYLE_KEYS = new Set(['color', 'backgroundColor', 'bold']);
const CONDITIONAL_COLOR = /^#?(?:[\da-f]{6}|[\da-f]{8})$/iu;
const CONDITIONAL_FORMULA_FUNCTIONS = new Set(
  BUILTIN_FORMULA_COMPATIBILITY.map(({ name }) => name),
);

function hasOnlySupportedFormulaFunctions(ast: FormulaAst): boolean {
  if (ast.kind === 'call') {
    return (
      CONDITIONAL_FORMULA_FUNCTIONS.has(ast.name) &&
      ast.arguments.every(hasOnlySupportedFormulaFunctions)
    );
  }
  if (ast.kind === 'unary') return hasOnlySupportedFormulaFunctions(ast.operand);
  if (ast.kind === 'binary') {
    return (
      hasOnlySupportedFormulaFunctions(ast.left) && hasOnlySupportedFormulaFunctions(ast.right)
    );
  }
  return true;
}

function conditionalColorAt(value: unknown, path: string, context: ParseContext): string {
  const color = stringAt(value, path, context);
  if (!CONDITIONAL_COLOR.test(color)) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a six- or eight-digit RGB color`,
    );
  }
  return color;
}

function conditionalStyleAt(value: unknown, path: string, context: ParseContext): ConditionalStyle {
  const source = recordAt(value, path, context);
  for (const key of Object.keys(source ?? {})) {
    if (!CONDITIONAL_STYLE_KEYS.has(key)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${path}.${key}`,
        `${path}.${key} is not a supported conditional style field`,
      );
    }
  }
  return {
    ...(source?.color === undefined
      ? {}
      : { color: conditionalColorAt(source.color, `${path}.color`, context) }),
    ...(source?.backgroundColor === undefined
      ? {}
      : {
          backgroundColor: conditionalColorAt(
            source.backgroundColor,
            `${path}.backgroundColor`,
            context,
          ),
        }),
    ...(source?.bold === undefined
      ? {}
      : { bold: booleanAt(source.bold, `${path}.bold`, context) }),
  };
}

function conditionalFormulaAt(value: unknown, path: string, context: ParseContext): string {
  const formula = stringAt(value, path, context);
  if (formula.length === 0 || formula.startsWith('=')) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must be a non-empty formula without a leading equals sign`,
    );
    return formula;
  }
  try {
    if (!hasOnlySupportedFormulaFunctions(parseFormula(`=${formula}`))) {
      throw new TypeError('Unsupported conditional formula function');
    }
  } catch {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      path,
      `${path} must use valid restricted formula syntax`,
    );
  }
  return formula;
}

function conditionalFormattingAt(
  value: unknown,
  path: string,
  context: ParseContext,
): ConditionalFormat[] {
  if (value === undefined) return [];
  return arrayAt(value, path, context).map((entry, index): ConditionalFormat => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const range = rangeAt(source?.range, `${entryPath}.range`, context);
    const rangeRecord = recordAt(source?.range, `${entryPath}.range`, context);
    const qualifiedRange = {
      sheetId: stringAt(
        rangeRecord?.sheetId,
        `${entryPath}.range.sheetId`,
        context,
      ) as DocumentSheetId,
      ...range,
    };
    if (!isNormalized(range)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.range`,
        'Conditional formatting range must be normalized',
      );
    }
    if (source?.type === 'color-scale') {
      return {
        type: 'color-scale',
        range: qualifiedRange,
        minimumColor: conditionalColorAt(source.minimumColor, `${entryPath}.minimumColor`, context),
        ...(source.midpointColor === undefined
          ? {}
          : {
              midpointColor: conditionalColorAt(
                source.midpointColor,
                `${entryPath}.midpointColor`,
                context,
              ),
            }),
        maximumColor: conditionalColorAt(source.maximumColor, `${entryPath}.maximumColor`, context),
      };
    }
    if (source?.type !== 'cell-is') {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.type`,
        `${entryPath}.type must be color-scale or cell-is`,
      );
    }
    const operator =
      typeof source?.operator === 'string' && CONDITIONAL_OPERATORS.has(source.operator)
        ? source.operator
        : 'equal';
    if (source?.operator !== operator) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.operator`,
        `${entryPath}.operator is not supported`,
      );
    }
    const formula = conditionalFormulaAt(source?.formula, `${entryPath}.formula`, context);
    const formula2 =
      source?.formula2 === undefined
        ? undefined
        : conditionalFormulaAt(source.formula2, `${entryPath}.formula2`, context);
    const needsSecond = operator === 'between' || operator === 'notBetween';
    if ((needsSecond && formula2 === undefined) || (!needsSecond && formula2 !== undefined)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.formula2`,
        `${entryPath}.formula2 does not match the conditional operator`,
      );
    }
    return {
      type: 'cell-is',
      range: qualifiedRange,
      operator: operator as Extract<ConditionalFormat, { readonly type: 'cell-is' }>['operator'],
      formula,
      ...(formula2 === undefined ? {} : { formula2 }),
      style: conditionalStyleAt(source?.style, `${entryPath}.style`, context),
    };
  });
}

const FILTER_VIEW_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
  'contains',
]);

function filterViewsAt(value: unknown, path: string, context: ParseContext): FilterView[] {
  if (value === undefined) return [];
  return arrayAt(value, path, context).map((entry, index): FilterView => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const rangeRecord = recordAt(source?.range, `${entryPath}.range`, context);
    const range = rangeAt(source?.range, `${entryPath}.range`, context);
    const sorts = arrayAt(source?.sorts, `${entryPath}.sorts`, context).map((sort, sortIndex) => {
      const sortPath = `${entryPath}.sorts[${sortIndex}]`;
      const item = recordAt(sort, sortPath, context);
      const direction =
        item?.direction === 'descending' || item?.direction === 'ascending'
          ? item.direction
          : 'ascending';
      if (item?.direction !== direction) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${sortPath}.direction`,
          `${sortPath}.direction must be ascending or descending`,
        );
      }
      return {
        column: indexAt(item?.column, `${sortPath}.column`, context),
        direction: direction as 'ascending' | 'descending',
      };
    });
    const filters = arrayAt(source?.filters, `${entryPath}.filters`, context).map(
      (filter, filterIndex): FilterViewPredicate => {
        const filterPath = `${entryPath}.filters[${filterIndex}]`;
        const item = recordAt(filter, filterPath, context);
        const operator =
          typeof item?.operator === 'string' && FILTER_VIEW_OPERATORS.has(item.operator)
            ? item.operator
            : 'equal';
        if (item?.operator !== operator) {
          addDiagnostic(
            context,
            'DOCUMENT_SCHEMA_INVALID',
            `${filterPath}.operator`,
            `${filterPath}.operator is unsupported`,
          );
        }
        const scalar = item?.value;
        if (
          typeof scalar !== 'string' &&
          typeof scalar !== 'boolean' &&
          (typeof scalar !== 'number' || !Number.isFinite(scalar))
        ) {
          addDiagnostic(
            context,
            'DOCUMENT_SCHEMA_INVALID',
            `${filterPath}.value`,
            `${filterPath}.value must be a finite scalar`,
          );
        }
        return {
          column: indexAt(item?.column, `${filterPath}.column`, context),
          operator: operator as FilterViewPredicate['operator'],
          value:
            typeof scalar === 'string' || typeof scalar === 'boolean'
              ? scalar
              : finiteAt(scalar, `${filterPath}.value`, context),
        };
      },
    );
    if (source?.visibility !== 'document') {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.visibility`,
        `${entryPath}.visibility must be document`,
      );
    }
    return {
      id: stringAt(source?.id, `${entryPath}.id`, context),
      name: displayStringAt(source?.name, `${entryPath}.name`, context),
      range: {
        sheetId: stringAt(
          rangeRecord?.sheetId,
          `${entryPath}.range.sheetId`,
          context,
        ) as DocumentSheetId,
        ...range,
      },
      sorts,
      filters,
      visibility: 'document',
    };
  });
}

function objectOffsetAt(
  value: unknown,
  path: string,
  context: ParseContext,
): { x: number; y: number } {
  const source = recordAt(value, path, context);
  return {
    x: finiteAt(source?.x, `${path}.x`, context),
    y: finiteAt(source?.y, `${path}.y`, context),
  };
}

function objectAddressAt(
  value: unknown,
  path: string,
  context: ParseContext,
): { sheetId: DocumentSheetId; row: number; column: number } {
  const source = recordAt(value, path, context);
  return {
    sheetId: stringAt(source?.sheetId, `${path}.sheetId`, context) as DocumentSheetId,
    row: indexAt(source?.row, `${path}.row`, context),
    column: indexAt(source?.column, `${path}.column`, context),
  };
}

function objectAnchorAt(value: unknown, path: string, context: ParseContext): ObjectAnchor {
  const source = recordAt(value, path, context);
  if (source?.type === 'absolute') {
    const rect = recordAt(source.rect, `${path}.rect`, context);
    return {
      type: 'absolute',
      rect: {
        x: finiteAt(rect?.x, `${path}.rect.x`, context),
        y: finiteAt(rect?.y, `${path}.rect.y`, context),
        width: nonNegativeFiniteAt(rect?.width, `${path}.rect.width`, context),
        height: nonNegativeFiniteAt(rect?.height, `${path}.rect.height`, context),
      },
    };
  }
  if (source?.type === 'one-cell') {
    const size = recordAt(source.size, `${path}.size`, context);
    return {
      type: 'one-cell',
      cell: objectAddressAt(source.cell, `${path}.cell`, context),
      offset: objectOffsetAt(source.offset, `${path}.offset`, context),
      size: {
        width: nonNegativeFiniteAt(size?.width, `${path}.size.width`, context),
        height: nonNegativeFiniteAt(size?.height, `${path}.size.height`, context),
      },
    };
  }
  if (source?.type !== 'two-cell') {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      `${path}.type`,
      `${path}.type must be absolute, one-cell, or two-cell`,
    );
  }
  const from = recordAt(source?.from, `${path}.from`, context);
  const to = recordAt(source?.to, `${path}.to`, context);
  return {
    type: 'two-cell',
    from: {
      ...objectAddressAt(from, `${path}.from`, context),
      offset: objectOffsetAt(from?.offset, `${path}.from.offset`, context),
    },
    to: {
      ...objectAddressAt(to, `${path}.to`, context),
      offset: objectOffsetAt(to?.offset, `${path}.to.offset`, context),
    },
  };
}

function sheetObjectsAt(value: unknown, path: string, context: ParseContext): SheetObject[] {
  if (value === undefined) return [];
  return arrayAt(value, path, context).map((entry, index): SheetObject => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const repeat =
      source?.templateRepeat === 'per-item' ||
      source?.templateRepeat === 'shared' ||
      source?.templateRepeat === 'forbidden'
        ? source.templateRepeat
        : 'forbidden';
    if (source?.templateRepeat !== repeat) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.templateRepeat`,
        `${entryPath}.templateRepeat is invalid`,
      );
    }
    const accessibility = recordAt(source?.accessibility, `${entryPath}.accessibility`, context);
    const rotation =
      source?.rotation === undefined
        ? undefined
        : ((finiteAt(source.rotation, `${entryPath}.rotation`, context) % 360) + 360) % 360;
    const common = {
      id: stringAt(source?.id, `${entryPath}.id`, context) as ObjectId,
      anchor: objectAnchorAt(source?.anchor, `${entryPath}.anchor`, context),
      zIndex: indexAt(source?.zIndex, `${entryPath}.zIndex`, context),
      locked: booleanAt(source?.locked, `${entryPath}.locked`, context),
      templateRepeat: repeat,
      ...(rotation === undefined ? {} : { rotation }),
      accessibility: {
        name: displayStringAt(accessibility?.name, `${entryPath}.accessibility.name`, context),
        ...(accessibility?.description === undefined
          ? {}
          : {
              description: displayStringAt(
                accessibility.description,
                `${entryPath}.accessibility.description`,
                context,
              ),
            }),
      },
    } as const;
    if (source?.kind === 'image') {
      const fit =
        source.fit === undefined ||
        source.fit === 'contain' ||
        source.fit === 'cover' ||
        source.fit === 'fill'
          ? source.fit
          : undefined;
      if (source.fit !== fit) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${entryPath}.fit`,
          `${entryPath}.fit is invalid`,
        );
      }
      return {
        ...common,
        kind: 'image',
        resourceId: stringAt(source.resourceId, `${entryPath}.resourceId`, context) as ResourceId,
        ...(fit === undefined ? {} : { fit }),
      };
    }
    if (source?.kind === 'shape') {
      const shape =
        source.shape === 'rectangle' || source.shape === 'ellipse' || source.shape === 'line'
          ? source.shape
          : 'rectangle';
      if (source.shape !== shape) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${entryPath}.shape`,
          `${entryPath}.shape must be rectangle, ellipse, or line`,
        );
      }
      const style = recordAt(source.style, `${entryPath}.style`, context);
      return {
        ...common,
        kind: 'shape',
        shape,
        style: {
          ...(style?.fill === undefined
            ? {}
            : { fill: stringAt(style.fill, `${entryPath}.style.fill`, context) }),
          ...(style?.stroke === undefined
            ? {}
            : { stroke: stringAt(style.stroke, `${entryPath}.style.stroke`, context) }),
          ...(style?.strokeWidth === undefined
            ? {}
            : {
                strokeWidth: nonNegativeFiniteAt(
                  style.strokeWidth,
                  `${entryPath}.style.strokeWidth`,
                  context,
                ),
              }),
        },
      };
    }
    if (source?.kind !== 'text-box') {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.kind`,
        `${entryPath}.kind must be image, shape, or text-box`,
      );
    }
    const style = recordAt(source?.style, `${entryPath}.style`, context);
    const align =
      style?.horizontalAlign === undefined ||
      style.horizontalAlign === 'left' ||
      style.horizontalAlign === 'center' ||
      style.horizontalAlign === 'right'
        ? style?.horizontalAlign
        : undefined;
    return {
      ...common,
      kind: 'text-box',
      text: displayStringAt(source?.text, `${entryPath}.text`, context),
      style: {
        color: stringAt(style?.color, `${entryPath}.style.color`, context),
        fontFamily: stringAt(style?.fontFamily, `${entryPath}.style.fontFamily`, context),
        fontSize: nonNegativeFiniteAt(style?.fontSize, `${entryPath}.style.fontSize`, context),
        ...(align === undefined ? {} : { horizontalAlign: align }),
      },
    };
  });
}

const TABLE_NAME_PATTERN = /^[A-Za-z_][A-Za-z0-9_.]*$/u;

function structuredTablesAt(
  value: unknown,
  path: string,
  context: ParseContext,
  ownerSheetId: DocumentSheetId,
  rowCount: number | undefined,
  columnCount: number | undefined,
): StructuredTable[] {
  if (value === undefined) return [];
  const tables = arrayAt(value, path, context).map((entry, index): StructuredTable => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const rangeRecord = recordAt(source?.range, `${entryPath}.range`, context);
    const range = rangeAt(source?.range, `${entryPath}.range`, context);
    const sheetId = stringAt(
      rangeRecord?.sheetId,
      `${entryPath}.range.sheetId`,
      context,
    ) as DocumentSheetId;
    const name = stringAt(source?.name, `${entryPath}.name`, context);
    if (!TABLE_NAME_PATTERN.test(name)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.name`,
        'Structured table names must start with a letter or underscore and contain only letters, digits, underscores, or periods',
      );
    }
    if (sheetId !== ownerSheetId) {
      addDiagnostic(
        context,
        'DANGLING_REFERENCE',
        `${entryPath}.range.sheetId`,
        'Structured table range must belong to its owning worksheet',
      );
    }
    if (
      !isNormalized(range) ||
      range.end.row <= range.start.row ||
      (rowCount !== undefined && range.end.row >= rowCount) ||
      (columnCount !== undefined && range.end.column >= columnCount)
    ) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${entryPath}.range`,
        'Structured table range must include a header and data row and remain inside the logical worksheet size',
      );
    }
    const columns = arrayAt(source?.columns, `${entryPath}.columns`, context).map(
      (column, columnIndex): StructuredTableColumn => {
        const columnPath = `${entryPath}.columns[${columnIndex}]`;
        const item = recordAt(column, columnPath, context);
        const dataType =
          item?.dataType === 'text' ||
          item?.dataType === 'number' ||
          item?.dataType === 'boolean' ||
          item?.dataType === 'date' ||
          item?.dataType === 'mixed'
            ? item.dataType
            : undefined;
        if (item?.dataType !== undefined && dataType === undefined) {
          addDiagnostic(
            context,
            'DOCUMENT_SCHEMA_INVALID',
            `${columnPath}.dataType`,
            'Structured table column dataType is unsupported',
          );
        }
        return {
          id: stringAt(item?.id, `${columnPath}.id`, context) as TableColumnId,
          name: stringAt(item?.name, `${columnPath}.name`, context),
          ...(dataType === undefined ? {} : { dataType }),
        };
      },
    );
    if (columns.length !== range.end.column - range.start.column + 1) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${entryPath}.columns`,
        'Structured table column count must match its range width',
      );
    }
    duplicateDiagnostics(context, columns, `${entryPath}.columns`);
    const columnNames = new Set<string>();
    columns.forEach((column, columnIndex) => {
      const key = column.name.toLocaleLowerCase('en-US');
      if (columnNames.has(key)) {
        addDiagnostic(
          context,
          'DUPLICATE_ID',
          `${entryPath}.columns[${columnIndex}].name`,
          `Duplicate structured table column name ${column.name}`,
        );
      }
      columnNames.add(key);
    });
    if (source?.headerRows !== undefined && source.headerRows !== 1) {
      addDiagnostic(
        context,
        'INVALID_TABLE_HEADER',
        `${entryPath}.headerRows`,
        'Structured tables currently require exactly one header row',
      );
    }
    const tableFilter = filterAt(source?.filter, `${entryPath}.filter`, context);
    if (tableFilter?.range !== undefined) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.filter.range`,
        'Structured table filters inherit the table range',
      );
    }
    tableFilter?.filters.forEach((filter, filterIndex) => {
      if (filter.column < range.start.column || filter.column > range.end.column) {
        addDiagnostic(
          context,
          'INVALID_RANGE',
          `${entryPath}.filter.filters[${filterIndex}].column`,
          'Structured table filter column must be within the table range',
        );
      }
    });
    if (
      tableFilter?.sort !== undefined &&
      tableFilter.sort !== null &&
      (tableFilter.sort.column < range.start.column || tableFilter.sort.column > range.end.column)
    ) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${entryPath}.filter.sort.column`,
        'Structured table sort column must be within the table range',
      );
    }
    return {
      id: stringAt(source?.id, `${entryPath}.id`, context) as TableId,
      name,
      range: { sheetId, ...range },
      columns,
      ...(source?.headerRows === undefined ? {} : { headerRows: 1 as const }),
      ...(source?.totalsRow === undefined
        ? {}
        : { totalsRow: booleanAt(source.totalsRow, `${entryPath}.totalsRow`, context) }),
      ...(source?.style === undefined
        ? {}
        : { style: stringAt(source.style, `${entryPath}.style`, context) }),
      ...(source?.autoExpand === undefined
        ? {}
        : { autoExpand: booleanAt(source.autoExpand, `${entryPath}.autoExpand`, context) }),
      ...(tableFilter === undefined
        ? {}
        : {
            filter: {
              filters: tableFilter.filters,
              ...(tableFilter.sort === undefined ? {} : { sort: tableFilter.sort }),
            },
          }),
    };
  });
  duplicateDiagnostics(context, tables, path);
  tables.forEach((table, index) => {
    const conflict = tables.findIndex(
      (candidate, candidateIndex) =>
        candidateIndex < index && rangesOverlap(candidate.range, table.range),
    );
    if (conflict >= 0) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${path}[${index}].range`,
        `Structured table range overlaps ${tables[conflict]!.name}`,
      );
    }
  });
  const dataRows = (table: StructuredTable): { start: number; end: number } => ({
    start: table.range.start.row + (table.headerRows ?? 1),
    end: table.range.end.row - (table.totalsRow === true ? 1 : 0),
  });
  tables.forEach((table, index) => {
    if (!hasActiveStructuredTableProjection(table)) return;
    const rows = dataRows(table);
    const conflict = tables.findIndex((candidate, candidateIndex) => {
      if (candidateIndex >= index || !hasActiveStructuredTableProjection(candidate)) return false;
      const candidateRows = dataRows(candidate);
      return rows.start <= candidateRows.end && candidateRows.start <= rows.end;
    });
    if (conflict >= 0) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${path}[${index}].filter`,
        `Active table row projections overlap ${tables[conflict]!.name}`,
      );
    }
  });
  return tables;
}

function analysisRangeAt(value: unknown, path: string, context: ParseContext) {
  const source = recordAt(value, path, context);
  const range = rangeAt(value, path, context);
  const sheetId = stringAt(source?.sheetId, `${path}.sheetId`, context) as DocumentSheetId;
  if (!isNormalized(range)) {
    addDiagnostic(context, 'INVALID_RANGE', path, 'Analysis range must be normalized');
  }
  return { sheetId, ...range };
}

function chartsAt(value: unknown, path: string, context: ParseContext): ChartDefinition[] {
  if (value === undefined) return [];
  const supported = new Set<ChartType>([
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'scatter',
    'combo',
  ]);
  const charts = arrayAt(value, path, context).map((entry, index): ChartDefinition => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const type = source?.type as ChartType;
    if (!supported.has(type)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.type`,
        'Chart type is unsupported',
      );
    }
    const series = arrayAt(source?.series, `${entryPath}.series`, context).map(
      (item, seriesIndex): ChartSeriesDefinition => {
        const seriesPath = `${entryPath}.series[${seriesIndex}]`;
        const record = recordAt(item, seriesPath, context);
        return {
          id: stringAt(record?.id, `${seriesPath}.id`, context),
          ...(record?.name === undefined
            ? {}
            : { name: displayStringAt(record.name, `${seriesPath}.name`, context) }),
          values: analysisRangeAt(record?.values, `${seriesPath}.values`, context),
        };
      },
    );
    duplicateDiagnostics(context, series, `${entryPath}.series`);
    const repeat =
      source?.templateRepeat === 'shared' ||
      source?.templateRepeat === 'per-item' ||
      source?.templateRepeat === 'forbidden'
        ? source.templateRepeat
        : undefined;
    if (source?.templateRepeat !== undefined && repeat === undefined) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.templateRepeat`,
        'Chart templateRepeat is unsupported',
      );
    }
    return {
      id: stringAt(source?.id, `${entryPath}.id`, context),
      type: supported.has(type) ? type : 'column',
      ...(source?.title === undefined
        ? {}
        : { title: displayStringAt(source.title, `${entryPath}.title`, context) }),
      ...(source?.categories === undefined
        ? {}
        : {
            categories: analysisRangeAt(source.categories, `${entryPath}.categories`, context),
          }),
      series,
      ...(source?.anchor === undefined
        ? {}
        : { anchor: objectAnchorAt(source.anchor, `${entryPath}.anchor`, context) }),
      ...(repeat === undefined ? {} : { templateRepeat: repeat }),
    };
  });
  duplicateDiagnostics(context, charts, path);
  return charts;
}

function sparklinesAt(value: unknown, path: string, context: ParseContext): SparklineDefinition[] {
  if (value === undefined) return [];
  const supported = new Set<SparklineType>(['line', 'column', 'win-loss']);
  const sparklines = arrayAt(value, path, context).map((entry, index): SparklineDefinition => {
    const entryPath = `${path}[${index}]`;
    const source = recordAt(entry, entryPath, context);
    const type = source?.type as SparklineType;
    if (!supported.has(type)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${entryPath}.type`,
        'Sparkline type is unsupported',
      );
    }
    const target = recordAt(source?.target, `${entryPath}.target`, context);
    return {
      id: stringAt(source?.id, `${entryPath}.id`, context),
      type: supported.has(type) ? type : 'line',
      source: analysisRangeAt(source?.source, `${entryPath}.source`, context),
      target: {
        sheetId: stringAt(
          target?.sheetId,
          `${entryPath}.target.sheetId`,
          context,
        ) as DocumentSheetId,
        ...pointAt(source?.target, `${entryPath}.target`, context),
      },
      ...(source?.color === undefined
        ? {}
        : { color: stringAt(source.color, `${entryPath}.color`, context) }),
      ...(source?.negativeColor === undefined
        ? {}
        : {
            negativeColor: stringAt(source.negativeColor, `${entryPath}.negativeColor`, context),
          }),
    };
  });
  duplicateDiagnostics(context, sparklines, path);
  return sparklines;
}

function sheetAt(value: unknown, path: string, context: ParseContext): Sheet {
  const record = recordAt(value, path, context);
  const merges = arrayAt(record?.merges, `${path}.merges`, context).map((merge, index) =>
    rangeAt(merge, `${path}.merges[${index}]`, context),
  );
  const cells = arrayAt(record?.cells, `${path}.cells`, context).map((entry, index): SparseCell => {
    const entryPath = `${path}.cells[${index}]`;
    const item = recordAt(entry, entryPath, context);
    return {
      row: indexAt(item?.row, `${entryPath}.row`, context),
      column: indexAt(item?.column, `${entryPath}.column`, context),
      cell: cellAt(item?.cell, `${entryPath}.cell`, context),
    };
  });
  const rows = layoutAt<SheetRow>(record?.rows, `${path}.rows`, context, 'height');
  const columns = layoutAt<SheetColumn>(record?.columns, `${path}.columns`, context, 'width');
  const rowCount =
    record?.rowCount === undefined
      ? undefined
      : indexAt(record.rowCount, `${path}.rowCount`, context);
  const columnCount =
    record?.columnCount === undefined
      ? undefined
      : indexAt(record.columnCount, `${path}.columnCount`, context);
  const groups = groupsAt(record?.groups, `${path}.groups`, context, rowCount, columnCount);
  const freeze =
    record?.freeze === undefined ? undefined : pointAt(record.freeze, `${path}.freeze`, context);
  const filter = filterAt(record?.filter, `${path}.filter`, context);
  const visibility =
    record?.visibility === undefined ||
    record.visibility === 'visible' ||
    record.visibility === 'hidden' ||
    record.visibility === 'very-hidden'
      ? (record?.visibility ?? 'visible')
      : 'visible';
  if (record?.visibility !== undefined && record.visibility !== visibility) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      `${path}.visibility`,
      `${path}.visibility must be visible, hidden, or very-hidden`,
    );
  }
  const conditionalFormatting = conditionalFormattingAt(
    record?.conditionalFormatting,
    `${path}.conditionalFormatting`,
    context,
  );
  const filterViews = filterViewsAt(record?.filterViews, `${path}.filterViews`, context);
  const objects = sheetObjectsAt(record?.objects, `${path}.objects`, context);
  const sheetId = stringAt(record?.id, `${path}.id`, context) as DocumentSheetId;
  const tables = structuredTablesAt(
    record?.tables,
    `${path}.tables`,
    context,
    sheetId,
    rowCount,
    columnCount,
  );
  const charts = chartsAt(record?.charts, `${path}.charts`, context);
  const sparklines = sparklinesAt(record?.sparklines, `${path}.sparklines`, context);

  for (const [index, merge] of merges.entries()) {
    if (!isNormalized(merge)) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `${path}.merges[${index}]`,
        'Merge range must be normalized',
      );
    }
  }
  const seenCoordinates = new Set<string>();
  cells.forEach(({ row, column }, index) => {
    const coordinate = `${row}:${column}`;
    if (seenCoordinates.has(coordinate)) {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        `${path}.cells[${index}]`,
        'Sparse cell coordinates must be unique',
      );
    }
    seenCoordinates.add(coordinate);
  });

  return {
    id: sheetId,
    name: displayStringAt(record?.name, `${path}.name`, context),
    cells,
    merges,
    ...(rowCount === undefined ? {} : { rowCount }),
    ...(columnCount === undefined ? {} : { columnCount }),
    rows,
    columns,
    groups,
    ...(freeze === undefined ? {} : { freeze }),
    ...(filter === undefined ? {} : { filter }),
    visibility,
    conditionalFormatting,
    filterViews,
    objects,
    tables,
    charts,
    sparklines,
  };
}

function templateAt(
  value: unknown,
  path: string,
  context: ParseContext,
): StoredSpreadsheetTemplate {
  const record = recordAt(value, path, context);
  const bindings = arrayAt(record?.bindings, `${path}.bindings`, context).map(
    (binding, index) =>
      canonicalJson(
        jsonAt(binding, `${path}.bindings[${index}]`, context),
      ) as unknown as TemplateBinding,
  );
  const printProfiles = arrayAt(record?.printProfiles, `${path}.printProfiles`, context).map(
    (profile, index) =>
      canonicalJson(
        jsonAt(profile, `${path}.printProfiles[${index}]`, context),
      ) as unknown as TemplatePrintProfile,
  );
  return {
    id: stringAt(record?.id, `${path}.id`, context) as TemplateId,
    name: stringAt(record?.name, `${path}.name`, context),
    bindings,
    printProfiles,
  };
}

function resourceAt(value: unknown, path: string, context: ParseContext): ResourceMetadata {
  const record = recordAt(value, path, context);
  return {
    id: stringAt(record?.id, `${path}.id`, context) as ResourceId,
    kind: stringAt(record?.kind, `${path}.kind`, context),
    ...(record?.mimeType === undefined
      ? {}
      : { mimeType: stringAt(record.mimeType, `${path}.mimeType`, context) }),
    ...(record?.url === undefined ? {} : { url: stringAt(record.url, `${path}.url`, context) }),
    ...(record?.byteLength === undefined
      ? {}
      : { byteLength: indexAt(record.byteLength, `${path}.byteLength`, context) }),
    ...(record?.metadata === undefined
      ? {}
      : { metadata: canonicalJson(jsonAt(record.metadata, `${path}.metadata`, context)) }),
  };
}

function duplicateDiagnostics(
  context: ParseContext,
  values: readonly { readonly id: string }[],
  path: string,
): void {
  const seen = new Set<string>();
  values.forEach(({ id }, index) => {
    if (seen.has(id)) {
      addDiagnostic(context, 'DUPLICATE_ID', `${path}[${index}].id`, `Duplicate stable ID ${id}`);
    }
    seen.add(id);
  });
}

function rangesOverlap(left: SheetRange, right: SheetRange): boolean {
  return (
    left.start.row <= right.end.row &&
    right.start.row <= left.end.row &&
    left.start.column <= right.end.column &&
    right.start.column <= left.end.column
  );
}

function formulaSheetTokens(ast: FormulaAst): readonly string[] {
  if (ast.kind === 'reference') {
    return ast.reference.sheetToken === undefined ? [] : [ast.reference.sheetToken];
  }
  if (ast.kind === 'range') {
    return [ast.start.sheetToken, ast.end.sheetToken].filter(
      (token): token is string => token !== undefined,
    );
  }
  if (ast.kind === 'unary') return formulaSheetTokens(ast.operand);
  if (ast.kind === 'binary') {
    return [...formulaSheetTokens(ast.left), ...formulaSheetTokens(ast.right)];
  }
  if (ast.kind === 'call') return ast.arguments.flatMap(formulaSheetTokens);
  return [];
}

function validateReferences(document: SpreadsheetDocument, context: ParseContext): void {
  const sheetIds = new Set(document.workbook.sheets.map(({ id }) => id));
  const styleIds = new Set(document.workbook.styles.map(({ id }) => id));
  const validationIds = new Set(document.workbook.validations.map(({ id }) => id));
  const resourceIds = new Set(document.resources.items.map(({ id }) => id));
  const templateIds = new Set(document.templates.map(({ id }) => id));
  const tableIds = new Set<string>();
  const tableColumnIds = new Set<string>();
  const tableNames = new Set<string>();

  document.workbook.sheets.forEach((sheet, sheetIndex) => {
    sheet.tables.forEach((table, tableIndex) => {
      const tablePath = `$.workbook.sheets[${sheetIndex}].tables[${tableIndex}]`;
      if (tableIds.has(table.id)) {
        addDiagnostic(
          context,
          'DUPLICATE_ID',
          `${tablePath}.id`,
          `Duplicate stable ID ${table.id}`,
        );
      }
      tableIds.add(table.id);
      table.columns.forEach((column, columnIndex) => {
        if (tableColumnIds.has(column.id)) {
          addDiagnostic(
            context,
            'DUPLICATE_ID',
            `${tablePath}.columns[${columnIndex}].id`,
            `Duplicate stable ID ${column.id}`,
          );
        }
        tableColumnIds.add(column.id);
      });
      const nameKey = table.name.toLocaleLowerCase('en-US');
      if (tableNames.has(nameKey)) {
        addDiagnostic(
          context,
          'DUPLICATE_ID',
          `${tablePath}.name`,
          `Duplicate structured table name ${table.name}`,
        );
      }
      tableNames.add(nameKey);
    });
    duplicateDiagnostics(
      context,
      sheet.filterViews,
      `$.workbook.sheets[${sheetIndex}].filterViews`,
    );
    duplicateDiagnostics(context, sheet.objects, `$.workbook.sheets[${sheetIndex}].objects`);
    sheet.filterViews.forEach((view, viewIndex) => {
      const base = `$.workbook.sheets[${sheetIndex}].filterViews[${viewIndex}]`;
      if (view.range.sheetId !== sheet.id || !sheetIds.has(view.range.sheetId)) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${base}.range.sheetId`,
          'Saved-view range must reference its owning sheet',
        );
      }
      if (!isNormalized(view.range)) {
        addDiagnostic(
          context,
          'INVALID_RANGE',
          `${base}.range`,
          'Saved-view range must be normalized',
        );
      }
    });
    sheet.objects.forEach((object, objectIndex) => {
      const base = `$.workbook.sheets[${sheetIndex}].objects[${objectIndex}]`;
      const anchors =
        object.anchor.type === 'absolute'
          ? []
          : object.anchor.type === 'one-cell'
            ? [object.anchor.cell]
            : [object.anchor.from, object.anchor.to];
      if (anchors.some(({ sheetId }) => sheetId !== sheet.id || !sheetIds.has(sheetId))) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${base}.anchor`,
          'Object anchors must reference their owning sheet',
        );
      }
      if (
        object.anchor.type === 'two-cell' &&
        (object.anchor.from.row > object.anchor.to.row ||
          (object.anchor.from.row === object.anchor.to.row &&
            object.anchor.from.offset.y > object.anchor.to.offset.y) ||
          object.anchor.from.column > object.anchor.to.column ||
          (object.anchor.from.column === object.anchor.to.column &&
            object.anchor.from.offset.x > object.anchor.to.offset.x))
      ) {
        addDiagnostic(
          context,
          'OBJECT_ANCHOR_INVALID',
          `${base}.anchor`,
          'Two-cell object markers must not cross',
        );
      }
      if (object.kind === 'image' && !resourceIds.has(object.resourceId)) {
        addDiagnostic(
          context,
          'DANGLING_REFERENCE',
          `${base}.resourceId`,
          'Referenced object resourceId does not exist',
        );
      }
    });
    sheet.charts.forEach((chart, chartIndex) => {
      if (chart.anchor === undefined || chart.anchor.type === 'absolute') return;
      const base = `$.workbook.sheets[${sheetIndex}].charts[${chartIndex}].anchor`;
      const markers =
        chart.anchor.type === 'one-cell'
          ? [chart.anchor.cell]
          : [chart.anchor.from, chart.anchor.to];
      if (markers.some(({ sheetId }) => sheetId !== sheet.id || !sheetIds.has(sheetId))) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          base,
          'Chart anchors must reference their owning sheet',
        );
      }
      if (
        chart.anchor.type === 'two-cell' &&
        (chart.anchor.from.row > chart.anchor.to.row ||
          (chart.anchor.from.row === chart.anchor.to.row &&
            chart.anchor.from.offset.y > chart.anchor.to.offset.y) ||
          chart.anchor.from.column > chart.anchor.to.column ||
          (chart.anchor.from.column === chart.anchor.to.column &&
            chart.anchor.from.offset.x > chart.anchor.to.offset.x))
      ) {
        addDiagnostic(
          context,
          'OBJECT_ANCHOR_INVALID',
          base,
          'Two-cell chart markers must not cross',
        );
      }
    });
    sheet.conditionalFormatting.forEach((format, formatIndex) => {
      const base = `$.workbook.sheets[${sheetIndex}].conditionalFormatting[${formatIndex}]`;
      if (format.range.sheetId !== sheet.id || !sheetIds.has(format.range.sheetId)) {
        addDiagnostic(
          context,
          'DOCUMENT_SCHEMA_INVALID',
          `${base}.range.sheetId`,
          'Conditional formatting range must reference its owning sheet',
        );
      }
      if (format.type === 'cell-is') {
        for (const [field, source] of [
          ['formula', format.formula],
          ...(format.formula2 === undefined ? [] : ([['formula2', format.formula2]] as const)),
        ] as const) {
          try {
            const tokens = formulaSheetTokens(parseFormula(`=${source}`));
            if (
              tokens.some(
                (token) =>
                  document.workbook.sheets.filter(
                    ({ name }) => name.toLowerCase() === token.toLowerCase(),
                  ).length !== 1,
              )
            ) {
              addDiagnostic(
                context,
                'DOCUMENT_SCHEMA_INVALID',
                `${base}.${field}`,
                'Conditional formula references an unknown or ambiguous sheet',
              );
            }
          } catch {
            // Syntax is diagnosed while decoding the rule.
          }
        }
      }
    });
    sheet.cells.forEach(({ cell }, cellIndex) => {
      const base = `$.workbook.sheets[${sheetIndex}].cells[${cellIndex}].cell`;
      const references: readonly [string | undefined, ReadonlySet<string>, string][] = [
        [cell.styleId, styleIds, 'styleId'],
        [cell.validationId, validationIds, 'validationId'],
        [cell.resourceId, resourceIds, 'resourceId'],
        [cell.templateId, templateIds, 'templateId'],
      ];
      for (const [id, ids, field] of references) {
        if (id !== undefined && !ids.has(id)) {
          addDiagnostic(
            context,
            'DANGLING_REFERENCE',
            `${base}.${field}`,
            `Referenced ${field} does not exist`,
          );
        }
      }
    });
    const analysisReferences = [
      ...sheet.charts.flatMap((chart) => [
        ...(chart.categories === undefined ? [] : [chart.categories]),
        ...chart.series.map(({ values }) => values),
      ]),
      ...sheet.sparklines.flatMap((sparkline) => [sparkline.source, sparkline.target]),
    ];
    analysisReferences.forEach((reference, referenceIndex) => {
      if (!sheetIds.has(reference.sheetId)) {
        addDiagnostic(
          context,
          'DANGLING_REFERENCE',
          `$.workbook.sheets[${sheetIndex}].analysis[${referenceIndex}].sheetId`,
          'Analysis definition references a sheet which does not exist',
        );
      }
    });
    sheet.sparklines.forEach((sparkline, sparklineIndex) => {
      if (sparkline.target.sheetId !== sheet.id) {
        addDiagnostic(
          context,
          'DANGLING_REFERENCE',
          `$.workbook.sheets[${sheetIndex}].sparklines[${sparklineIndex}].target.sheetId`,
          'Sparkline target must belong to its owning worksheet',
        );
      }
    });
    for (const [collectionName, collection] of [
      ['rows', sheet.rows],
      ['columns', sheet.columns],
    ] as const) {
      collection.forEach((entry, entryIndex) => {
        if (entry.styleId !== undefined && !styleIds.has(entry.styleId)) {
          addDiagnostic(
            context,
            'DANGLING_REFERENCE',
            `$.workbook.sheets[${sheetIndex}].${collectionName}[${entryIndex}].styleId`,
            'Referenced styleId does not exist',
          );
        }
      });
    }
    sheet.merges.forEach((merge, mergeIndex) => {
      for (let previous = 0; previous < mergeIndex; previous += 1) {
        if (rangesOverlap(sheet.merges[previous]!, merge)) {
          addDiagnostic(
            context,
            'INVALID_MERGE',
            `$.workbook.sheets[${sheetIndex}].merges[${mergeIndex}]`,
            'Merge ranges must not overlap',
          );
          break;
        }
      }
    });
  });

  if (
    document.workbook.sheets.length > 0 &&
    !document.workbook.sheets.some(({ visibility }) => visibility === 'visible')
  ) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      '$.workbook.sheets',
      'At least one worksheet must be visible',
    );
  }

  document.templates.forEach((template, index) => {
    const templateRanges = [
      ...template.bindings.flatMap((binding) => (binding.type === 'value' ? [] : [binding.range])),
      ...template.printProfiles.flatMap((profile) => [
        ...profile.targets.flatMap((target) =>
          target.type === 'sheet' ? [] : target.type === 'range' ? [target.range] : target.ranges,
        ),
        ...(profile.repeatRows === undefined ? [] : [profile.repeatRows]),
        ...(profile.repeatColumns === undefined ? [] : [profile.repeatColumns]),
      ]),
    ];
    if (templateRanges.some((range) => !isNormalized(range))) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `$.templates[${index}]`,
        'Template ranges must be normalized',
      );
    }
    const referencedSheetIds = [
      ...template.bindings.flatMap((binding) =>
        binding.type === 'value' ? [binding.target.sheetId] : [binding.range.sheetId],
      ),
      ...template.printProfiles.flatMap((profile) =>
        profile.targets.flatMap((target) =>
          target.type === 'sheet'
            ? [target.sheetId]
            : target.type === 'range'
              ? [target.range.sheetId]
              : target.ranges.map(({ sheetId }) => sheetId),
        ),
      ),
    ];
    if (referencedSheetIds.some((sheetId) => !sheetIds.has(sheetId))) {
      addDiagnostic(
        context,
        'DANGLING_REFERENCE',
        `$.templates[${index}]`,
        'Template references a sheet which does not exist',
      );
    }
    template.bindings.forEach((binding, bindingIndex) => {
      if (
        binding.type !== 'repeat-rows' &&
        binding.type !== 'repeat-columns' &&
        binding.type !== 'repeat-range' &&
        binding.type !== 'repeat-page' &&
        binding.type !== 'repeat-sheet'
      ) {
        return;
      }
      const sheet = document.workbook.sheets.find(({ id }) => id === binding.range.sheetId);
      const conflict = sheet?.tables.find(
        (table) =>
          table.autoExpand === true &&
          table.range.end.row === binding.range.end.row &&
          table.range.start.column <= binding.range.end.column &&
          binding.range.start.column <= table.range.end.column,
      );
      if (conflict !== undefined) {
        addDiagnostic(
          context,
          'TABLE_TEMPLATE_BOUNDARY_CONFLICT',
          `$.templates[${index}].bindings[${bindingIndex}].range`,
          `Template repeat and structured table ${conflict.name} cannot both control the same expansion boundary`,
        );
      }
    });
  });
}

function decodeDocument(value: unknown, context: ParseContext): SpreadsheetDocument {
  const root = recordAt(value, '$', context);
  const workbook = recordAt(root?.workbook, '$.workbook', context);
  const settings = recordAt(workbook?.settings, '$.workbook.settings', context);
  const dateSystem =
    settings?.dateSystem === 'excel-1900' || settings?.dateSystem === 'excel-1904'
      ? settings.dateSystem
      : 'excel-1900';
  if (settings?.dateSystem !== dateSystem) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      '$.workbook.settings.dateSystem',
      'Date system must be excel-1900 or excel-1904',
      'document',
      'decode',
    );
  }

  const sheets = arrayAt(workbook?.sheets, '$.workbook.sheets', context).map((sheet, index) =>
    sheetAt(sheet, `$.workbook.sheets[${index}]`, context),
  );
  const styles = arrayAt(workbook?.styles, '$.workbook.styles', context).map((entry, index) => {
    const path = `$.workbook.styles[${index}]`;
    const item = recordAt(entry, path, context);
    return {
      id: stringAt(item?.id, `${path}.id`, context) as StyleId,
      value: canonicalJson(jsonAt(item?.value, `${path}.value`, context)),
    };
  });
  const validations = arrayAt(workbook?.validations, '$.workbook.validations', context).map(
    (entry, index) => {
      const path = `$.workbook.validations[${index}]`;
      const item = recordAt(entry, path, context);
      return {
        id: stringAt(item?.id, `${path}.id`, context) as ValidationId,
        value: canonicalJson(jsonAt(item?.value, `${path}.value`, context)),
      };
    },
  );
  const templates = arrayAt(root?.templates, '$.templates', context).map((template, index) =>
    templateAt(template, `$.templates[${index}]`, context),
  );
  const resourcesRecord = recordAt(root?.resources, '$.resources', context);
  const resources = arrayAt(resourcesRecord?.items, '$.resources.items', context).map(
    (resource, index) => resourceAt(resource, `$.resources.items[${index}]`, context),
  );
  const rawExtensions = recordAt(root?.extensions, '$.extensions', context);
  const extensions = Object.create(null) as Record<string, JsonValue>;
  for (const namespace of Object.keys(rawExtensions ?? {}).sort()) {
    if (!NAMESPACE_PATTERN.test(namespace)) {
      addDiagnostic(
        context,
        'INVALID_EXTENSION_DATA',
        `$.extensions.${namespace}`,
        'Extension keys must be namespaced',
        'extension',
      );
    }
    extensions[namespace] = jsonAt(
      rawExtensions?.[namespace],
      `$.extensions.${namespace}`,
      context,
      'INVALID_EXTENSION_DATA',
    );
  }

  const schemaVersion = root?.schemaVersion;
  if (schemaVersion !== 2) {
    addDiagnostic(
      context,
      typeof schemaVersion === 'number' ? 'UNSUPPORTED_SCHEMA_VERSION' : 'DOCUMENT_SCHEMA_INVALID',
      '$.schemaVersion',
      'Spreadsheet document schemaVersion must be exactly 2',
      'document',
      typeof schemaVersion === 'number' ? 'validate' : 'decode',
    );
  }

  const document: SpreadsheetDocument = {
    schemaVersion: 2,
    id: stringAt(root?.id, '$.id', context) as DocumentId,
    workbook: {
      sheets,
      styles,
      validations,
      settings: {
        dateSystem,
        ...(settings?.localeHint === undefined
          ? {}
          : {
              localeHint: stringAt(settings.localeHint, '$.workbook.settings.localeHint', context),
            }),
      },
    },
    templates,
    resources: { items: resources },
    extensions: extensions as ExtensionStore,
  };

  duplicateDiagnostics(context, sheets, '$.workbook.sheets');
  duplicateDiagnostics(context, styles, '$.workbook.styles');
  duplicateDiagnostics(context, validations, '$.workbook.validations');
  duplicateDiagnostics(context, templates, '$.templates');
  duplicateDiagnostics(context, resources, '$.resources.items');
  validateReferences(document, context);
  return document;
}

function canonicalizeDocument(document: SpreadsheetDocument): SpreadsheetDocument {
  return {
    ...document,
    workbook: {
      ...document.workbook,
      sheets: document.workbook.sheets.map((sheet) => ({
        ...sheet,
        cells: [...sheet.cells].sort(compareSparseCells),
        rows: [...sheet.rows].sort((left, right) => left.index - right.index),
        columns: [...sheet.columns].sort((left, right) => left.index - right.index),
        groups: [...sheet.groups],
        filterViews: [...sheet.filterViews].sort((left, right) =>
          compareCodeUnits(left.id, right.id),
        ),
        objects: [...sheet.objects].sort(
          (left, right) => left.zIndex - right.zIndex || compareCodeUnits(left.id, right.id),
        ),
        tables: [...sheet.tables].sort((left, right) => compareCodeUnits(left.id, right.id)),
        charts: [...sheet.charts].sort((left, right) => compareCodeUnits(left.id, right.id)),
        sparklines: [...sheet.sparklines].sort((left, right) =>
          compareCodeUnits(left.id, right.id),
        ),
      })),
      styles: [...document.workbook.styles].sort((left, right) =>
        compareCodeUnits(left.id, right.id),
      ),
      validations: [...document.workbook.validations].sort((left, right) =>
        compareCodeUnits(left.id, right.id),
      ),
    },
    resources: {
      items: [...document.resources.items].sort((left, right) =>
        compareCodeUnits(left.id, right.id),
      ),
    },
  };
}

/**
 * Parses and validates an unknown Workbook 2.0 input atomically.
 *
 * @param input - JSON text or an object-like document input.
 * @param options - Optional bounded-decoding limits.
 * @returns A frozen document on success or aggregated diagnostics on failure.
 */
export function parseSpreadsheetDocument(
  input: unknown,
  options: DocumentParseOptions = {},
): DocumentParseResult {
  const context: ParseContext = { diagnostics: [], activeJson: new WeakSet() };
  const limits = resolveLimits(options, context);
  if (limits === undefined) {
    return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
  }
  let decodedInput = input;

  if (typeof input === 'string') {
    if (new TextEncoder().encode(input).byteLength > limits.maxBytes) {
      addDiagnostic(
        context,
        'DOCUMENT_LIMIT_EXCEEDED',
        '$',
        '$ exceeds its configured document limit',
      );
      return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
    }
    try {
      decodedInput = JSON.parse(input) as unknown;
    } catch {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        '$',
        'Input must be valid JSON',
        'document',
        'decode',
      );
      return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
    }
  }

  try {
    decodedInput = deepFreeze(captureInput(decodedInput, limits));
  } catch (error) {
    if (error instanceof InputCaptureError) {
      addDiagnostic(
        context,
        error.code,
        error.path,
        error.message,
        error.code === 'INVALID_EXTENSION_DATA' ? 'extension' : 'document',
        error.code === 'DOCUMENT_LIMIT_EXCEEDED' ? 'validate' : 'decode',
      );
    } else {
      addDiagnostic(
        context,
        'DOCUMENT_SCHEMA_INVALID',
        '$',
        'Input could not be captured safely',
        'document',
        'decode',
      );
    }
    return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
  }

  const collectionLimitPath = exceedsCollectionLimits(decodedInput, limits);
  if (collectionLimitPath !== undefined) {
    addDiagnostic(
      context,
      'DOCUMENT_LIMIT_EXCEEDED',
      collectionLimitPath,
      `${collectionLimitPath} exceeds its configured document limit`,
    );
    return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
  }

  let document: SpreadsheetDocument;
  try {
    document = decodeDocument(decodedInput, context);
  } catch {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      '$',
      'Input could not be decoded safely',
      'document',
      'decode',
    );
    return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
  }
  if (context.diagnostics.length > 0) {
    return { ok: false, diagnostics: deepFreeze(context.diagnostics) };
  }
  return {
    ok: true,
    document: deepFreeze(canonicalizeDocument(document)),
    diagnostics: [],
  };
}
