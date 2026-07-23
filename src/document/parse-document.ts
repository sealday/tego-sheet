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
  DocumentCellRange,
  ExtensionStore,
  PrintProfile,
  ResourceMetadata,
  Sheet,
  SheetRange,
  SpreadsheetDocument,
  SpreadsheetTemplate,
} from './model/document';
import type {
  DocumentId,
  DocumentSheetId,
  ResourceId,
  StyleId,
  TemplateId,
  ValidationId,
} from './model/ids';
import { compareSparseCells } from './model/sparse-cells';
import type { SparseCell } from './model/sparse-cells';

/** Configurable safety limits enforced before deep document decoding. */
export interface DocumentLimits {
  /** Maximum number of sheets. */
  readonly maxSheets?: number;
  /** Maximum total number of sparse cells. */
  readonly maxCells?: number;
  /** Maximum total number of merge ranges. */
  readonly maxMerges?: number;
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
  maxBytes: 64 * 1024 * 1024,
} as const;
const NAMESPACE_PATTERN = /^[a-z0-9]+(?:[.-][a-z0-9]+)+$/i;

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
  readonly maxBytes: number;
}

const LIMIT_NAMES = ['maxSheets', 'maxCells', 'maxMerges', 'maxBytes'] as const;

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
    readonly code: 'DOCUMENT_SCHEMA_INVALID' | 'DOCUMENT_LIMIT_EXCEEDED',
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
}

function captureInput(input: unknown, limits: ResolvedDocumentLimits): unknown {
  const context: InputCaptureContext = {
    active: new WeakSet(),
    encoder: new TextEncoder(),
    limits,
    bytes: 0,
    cells: 0,
    merges: 0,
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
  const capture = (value: unknown, path: string, inArray: boolean): unknown => {
    if (
      value === null ||
      typeof value === 'string' ||
      typeof value === 'boolean' ||
      typeof value === 'number'
    ) {
      consume(JSON.stringify(value) ?? 'null');
      return value;
    }
    if (typeof value !== 'object') {
      if (inArray) consume('null');
      return value;
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
          output[index] = capture(descriptor.value, `${path}[${index}]`, true);
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
      if (prototype !== Object.prototype && prototype !== null) {
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
          value: capture(item, `${path}.${key}`, false),
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

  return capture(input, '$', false);
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
  }
  return undefined;
}

const isRecord = (value: unknown): value is Record<string, unknown> => {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
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
  };
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
    id: stringAt(record?.id, `${path}.id`, context) as DocumentSheetId,
    name: stringAt(record?.name, `${path}.name`, context),
    cells,
    merges,
  };
}

function printProfileAt(value: unknown, path: string, context: ParseContext): PrintProfile {
  const record = recordAt(value, path, context);
  const margins = recordAt(record?.margins, `${path}.margins`, context);
  const orientation =
    record?.orientation === 'portrait' || record?.orientation === 'landscape'
      ? record.orientation
      : 'portrait';
  if (record?.orientation !== orientation) {
    addDiagnostic(
      context,
      'DOCUMENT_SCHEMA_INVALID',
      `${path}.orientation`,
      `${path}.orientation must be portrait or landscape`,
      'document',
      'decode',
    );
  }
  return {
    paperSize: stringAt(record?.paperSize, `${path}.paperSize`, context),
    orientation,
    margins: {
      top: finiteAt(margins?.top, `${path}.margins.top`, context),
      right: finiteAt(margins?.right, `${path}.margins.right`, context),
      bottom: finiteAt(margins?.bottom, `${path}.margins.bottom`, context),
      left: finiteAt(margins?.left, `${path}.margins.left`, context),
    },
    ...(record?.scale === undefined
      ? {}
      : { scale: finiteAt(record.scale, `${path}.scale`, context) }),
  };
}

function templateAt(value: unknown, path: string, context: ParseContext): SpreadsheetTemplate {
  const record = recordAt(value, path, context);
  let range: DocumentCellRange | undefined;
  if (record?.range !== undefined) {
    const rangeRecord = recordAt(record.range, `${path}.range`, context);
    range = {
      sheetId: stringAt(rangeRecord?.sheetId, `${path}.range.sheetId`, context) as DocumentSheetId,
      ...rangeAt(record.range, `${path}.range`, context),
    };
    if (!isNormalized(range)) {
      addDiagnostic(context, 'INVALID_RANGE', `${path}.range`, 'Template range must be normalized');
    }
  }
  return {
    id: stringAt(record?.id, `${path}.id`, context) as TemplateId,
    name: stringAt(record?.name, `${path}.name`, context),
    sheetId: stringAt(record?.sheetId, `${path}.sheetId`, context) as DocumentSheetId,
    ...(range === undefined ? {} : { range }),
    printProfile: printProfileAt(record?.printProfile, `${path}.printProfile`, context),
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

function validateReferences(document: SpreadsheetDocument, context: ParseContext): void {
  const sheetIds = new Set(document.workbook.sheets.map(({ id }) => id));
  const styleIds = new Set(document.workbook.styles.map(({ id }) => id));
  const validationIds = new Set(document.workbook.validations.map(({ id }) => id));
  const resourceIds = new Set(document.resources.items.map(({ id }) => id));
  const templateIds = new Set(document.templates.map(({ id }) => id));

  document.workbook.sheets.forEach((sheet, sheetIndex) => {
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

  document.templates.forEach((template, index) => {
    if (!sheetIds.has(template.sheetId)) {
      addDiagnostic(
        context,
        'DANGLING_REFERENCE',
        `$.templates[${index}].sheetId`,
        'Referenced sheet does not exist',
      );
    }
    if (
      template.range !== undefined &&
      (template.range.sheetId !== template.sheetId || !sheetIds.has(template.range.sheetId))
    ) {
      addDiagnostic(
        context,
        'INVALID_RANGE',
        `$.templates[${index}].range`,
        'Template range must refer to its template sheet',
      );
    }
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
        'document',
        error.code === 'DOCUMENT_SCHEMA_INVALID' ? 'decode' : 'validate',
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
