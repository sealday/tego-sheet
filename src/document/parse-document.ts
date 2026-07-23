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

export interface DocumentLimits {
  readonly maxSheets?: number;
  readonly maxCells?: number;
  readonly maxBytes?: number;
}

export interface DocumentParseOptions {
  readonly limits?: DocumentLimits;
}

export type DocumentParseResult =
  | {
      readonly ok: true;
      readonly document: SpreadsheetDocument;
      readonly diagnostics: readonly [];
    }
  | {
      readonly ok: false;
      readonly diagnostics: readonly DocumentDiagnostic[];
    };

const DEFAULT_LIMITS = {
  maxSheets: 1_000,
  maxCells: 1_000_000,
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

export function parseSpreadsheetDocument(
  input: unknown,
  options: DocumentParseOptions = {},
): DocumentParseResult {
  const context: ParseContext = { diagnostics: [], activeJson: new WeakSet() };
  const limits = { ...DEFAULT_LIMITS, ...options.limits };
  let decodedInput = input;
  let encodedBytes: number | undefined;

  if (typeof input === 'string') {
    encodedBytes = new TextEncoder().encode(input).byteLength;
    if (encodedBytes > limits.maxBytes) {
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
  const cellCount = document.workbook.sheets.reduce(
    (total, sheet) => total + sheet.cells.length,
    0,
  );
  const measuredBytes =
    encodedBytes ?? new TextEncoder().encode(JSON.stringify(document)).byteLength;
  const limitChecks: readonly [number, number, string][] = [
    [document.workbook.sheets.length, limits.maxSheets, '$.workbook.sheets'],
    [cellCount, limits.maxCells, '$.workbook.sheets'],
    [measuredBytes, limits.maxBytes, '$'],
  ];
  for (const [actual, maximum, path] of limitChecks) {
    if (actual > maximum) {
      addDiagnostic(
        context,
        'DOCUMENT_LIMIT_EXCEEDED',
        path,
        `${path} exceeds its configured document limit`,
      );
    }
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
