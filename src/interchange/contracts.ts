import type { Diagnostic } from '../document';
import type { SpreadsheetDocument } from '../document';

export type InterchangeFormat = 'csv' | 'tsv' | 'xlsx' | 'ods';

export interface InterchangeLimits {
  readonly maxPackageBytes?: number;
  readonly maxUncompressedBytes?: number;
  readonly maxEntries?: number;
  readonly maxTextBytes?: number;
  readonly maxXmlBytes?: number;
  readonly maxFieldBytes?: number;
  readonly maxRows?: number;
  readonly maxColumns?: number;
  readonly maxCells?: number;
  readonly maxOutputBytes?: number;
}

export interface InterchangeReadOptions {
  readonly signal?: AbortSignal;
}

export interface DelimitedWriteOptions {
  readonly delimiter?: ',' | '\t';
  readonly lineEnding?: '\n' | '\r\n';
  readonly formulaInjectionProtection?: boolean;
  readonly signal?: AbortSignal;
  readonly maxOutputBytes?: number;
}

export interface InterchangeSecurityReport {
  readonly activeContentExecuted: false;
  readonly externalResourcesFetched: false;
  readonly warnings: readonly string[];
  readonly unsupportedFeatures: readonly string[];
}

export interface WorkbookImportResult {
  readonly format: InterchangeFormat;
  readonly document: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly security: InterchangeSecurityReport;
}

export type InterchangeInput = Uint8Array | ArrayBuffer | Blob | string;

export interface WorkbookReader {
  readonly format: InterchangeFormat;
  read(input: InterchangeInput, options?: InterchangeReadOptions): Promise<WorkbookImportResult>;
}

export interface WorkbookWriter {
  readonly format: 'csv' | 'tsv';
  write(document: SpreadsheetDocument, options?: DelimitedWriteOptions): Promise<Blob>;
}

export type InterchangeErrorCode =
  | 'ABORTED'
  | 'ACTIVE_CONTENT_REJECTED'
  | 'ARCHIVE_INVALID'
  | 'ARCHIVE_LIMIT_EXCEEDED'
  | 'CELL_LIMIT_EXCEEDED'
  | 'COLUMN_LIMIT_EXCEEDED'
  | 'DOCUMENT_INVALID'
  | 'EXTERNAL_RESOURCE_REJECTED'
  | 'FIELD_LIMIT_EXCEEDED'
  | 'MALFORMED_DELIMITED_TEXT'
  | 'MALFORMED_WORKBOOK'
  | 'OUTPUT_LIMIT_EXCEEDED'
  | 'ROW_LIMIT_EXCEEDED'
  | 'UNSUPPORTED_ARCHIVE_FEATURE'
  | 'XML_ENTITY_REJECTED'
  | 'XML_LIMIT_EXCEEDED';

export class InterchangeError extends Error {
  readonly name = 'InterchangeError';
  readonly security = securityReport();

  constructor(
    readonly code: InterchangeErrorCode,
    message: string,
    options?: { readonly cause?: unknown },
  ) {
    super(message, options);
  }
}

export const DEFAULT_INTERCHANGE_LIMITS = Object.freeze({
  maxPackageBytes: 64 * 1024 * 1024,
  maxUncompressedBytes: 128 * 1024 * 1024,
  maxEntries: 10_000,
  maxTextBytes: 64 * 1024 * 1024,
  maxXmlBytes: 32 * 1024 * 1024,
  maxFieldBytes: 1024 * 1024,
  maxRows: 1_048_576,
  maxColumns: 16_384,
  maxCells: 1_000_000,
  maxOutputBytes: 64 * 1024 * 1024,
});

export type ResolvedInterchangeLimits = Readonly<
  Required<{
    [Key in keyof InterchangeLimits]: number;
  }>
>;

export function resolveLimits(limits: InterchangeLimits = {}): ResolvedInterchangeLimits {
  const resolved = { ...DEFAULT_INTERCHANGE_LIMITS, ...limits };
  for (const [name, value] of Object.entries(resolved)) {
    if (!Number.isSafeInteger(value) || value < 0) {
      throw new TypeError(`${name} must be a non-negative safe integer`);
    }
  }
  return Object.freeze(resolved);
}

export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw new InterchangeError('ABORTED', 'Workbook interchange was aborted');
  }
}

export async function inputBytes(
  input: InterchangeInput,
  signal: AbortSignal | undefined,
): Promise<Uint8Array> {
  throwIfAborted(signal);
  let bytes: Uint8Array;
  if (typeof input === 'string') {
    bytes = new TextEncoder().encode(input);
  } else if (input instanceof Uint8Array) {
    bytes = input;
  } else if (input instanceof ArrayBuffer) {
    bytes = new Uint8Array(input);
  } else {
    bytes = new Uint8Array(await input.arrayBuffer());
  }
  throwIfAborted(signal);
  return bytes;
}

export function securityReport(
  warnings: readonly string[] = [],
  unsupportedFeatures: readonly string[] = [],
): InterchangeSecurityReport {
  return Object.freeze({
    activeContentExecuted: false,
    externalResourcesFetched: false,
    warnings: Object.freeze([...warnings]),
    unsupportedFeatures: Object.freeze([...new Set(unsupportedFeatures)]),
  });
}

export function importResult(
  format: InterchangeFormat,
  document: SpreadsheetDocument,
  warnings: readonly string[] = [],
  unsupportedFeatures: readonly string[] = [],
): WorkbookImportResult {
  return Object.freeze({
    format,
    document,
    diagnostics: Object.freeze([]),
    security: securityReport(warnings, unsupportedFeatures),
  });
}
