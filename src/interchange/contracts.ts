import type { Diagnostic } from '../document';
import type { SpreadsheetDocument } from '../document';

/** Supported spreadsheet interchange container or text formats. */
export type InterchangeFormat = 'csv' | 'tsv' | 'xlsx' | 'ods';

/** Resource limits enforced before and during workbook interchange. */
export interface InterchangeLimits {
  /** Maximum compressed package bytes. */
  readonly maxPackageBytes?: number;
  /** Maximum aggregate uncompressed archive bytes. */
  readonly maxUncompressedBytes?: number;
  /** Maximum archive entries. */
  readonly maxEntries?: number;
  /** Maximum delimited-text bytes. */
  readonly maxTextBytes?: number;
  /** Maximum bytes in one XML part. */
  readonly maxXmlBytes?: number;
  /** Maximum bytes in one delimited field. */
  readonly maxFieldBytes?: number;
  /** Maximum worksheet rows. */
  readonly maxRows?: number;
  /** Maximum worksheet columns. */
  readonly maxColumns?: number;
  /** Maximum non-empty cells. */
  readonly maxCells?: number;
  /** Maximum DrawingML objects materialized across the workbook. */
  readonly maxObjects?: number;
  /** Maximum unique embedded resources materialized across the workbook. */
  readonly maxResources?: number;
  /** Maximum aggregate bytes of unique embedded resources. */
  readonly maxResourceBytes?: number;
  /** Maximum emitted text bytes. */
  readonly maxOutputBytes?: number;
}

/** Per-read cancellation options. */
export interface InterchangeReadOptions {
  /** Cancels parsing before an atomic result is returned. */
  readonly signal?: AbortSignal;
}

/** Deterministic CSV and TSV writer options. */
export interface DelimitedWriteOptions {
  /** Field delimiter override. */
  readonly delimiter?: ',' | '\t';
  /** Output line ending. */
  readonly lineEnding?: '\n' | '\r\n';
  /** Prefixes spreadsheet-formula-like text to prevent injection. */
  readonly formulaInjectionProtection?: boolean;
  /** Cancels serialization. */
  readonly signal?: AbortSignal;
  /** Per-write output byte limit. */
  readonly maxOutputBytes?: number;
}

/** Per-write cancellation and output quota options shared by package writers. */
export interface InterchangeWriteOptions {
  /** Cancels serialization before an atomic result is returned. */
  readonly signal?: AbortSignal;
  /** Per-write finalized output byte limit. */
  readonly maxOutputBytes?: number;
}

/** Evidence that parsing performed no active or external execution. */
export interface InterchangeSecurityReport {
  /** Always false because readers never execute active content. */
  readonly activeContentExecuted: false;
  /** Always false because readers never fetch external resources. */
  readonly externalResourcesFetched: false;
  /** Recoverable security warnings. */
  readonly warnings: readonly string[];
  /** Recognized features that were not imported. */
  readonly unsupportedFeatures: readonly string[];
}

/** Immutable all-or-nothing workbook import result. */
export interface WorkbookImportResult {
  /** Detected or configured source format. */
  readonly format: InterchangeFormat;
  /** Fully parsed immutable document. */
  readonly document: SpreadsheetDocument;
  /** Structured import diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Security and degradation report. */
  readonly security: InterchangeSecurityReport;
}

/** Immutable structured result returned by semantic workbook writers. */
export interface WorkbookExportResult {
  /** Emitted workbook format. */
  readonly format: InterchangeFormat;
  /** Complete finalized output; never a partial package. */
  readonly blob: Blob;
  /** Structured semantic-degradation diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
}

/** Byte, blob, or text input accepted by workbook readers. */
export type InterchangeInput = Uint8Array | ArrayBuffer | Blob | string;

/** Atomic bounded workbook reader. */
export interface WorkbookReader {
  /** Format handled by this reader. */
  readonly format: InterchangeFormat;
  /** Reads one complete immutable document or rejects without partial output. */
  read(input: InterchangeInput, options?: InterchangeReadOptions): Promise<WorkbookImportResult>;
}

/** Atomic bounded workbook writer with a backward-compatible blob convenience method. */
export interface WorkbookWriter {
  /** Format emitted by this writer. */
  readonly format: InterchangeFormat;
  /** Serializes a document to a blob. */
  write(
    document: SpreadsheetDocument,
    options?: DelimitedWriteOptions | InterchangeWriteOptions,
  ): Promise<Blob>;
  /** Serializes a document and reports any intentional semantic degradation. */
  writeResult(
    document: SpreadsheetDocument,
    options?: DelimitedWriteOptions | InterchangeWriteOptions,
  ): Promise<WorkbookExportResult>;
}

/** Stable workbook interchange failure categories. */
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

/** Error thrown when interchange input violates syntax, safety, or resource limits. */
export class InterchangeError extends Error {
  /** Stable error class name. */
  readonly name = 'InterchangeError';
  /** Security report retained on rejected input. */
  readonly security = securityReport();

  /** Creates a stable interchange error. */
  constructor(
    /** Machine-readable failure category. */
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
  maxObjects: 10_000,
  maxResources: 10_000,
  maxResourceBytes: 64 * 1024 * 1024,
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
  maximumBytes: number,
): Promise<Uint8Array> {
  throwIfAborted(signal);
  const knownSize =
    typeof input === 'string'
      ? new TextEncoder().encode(input).byteLength
      : input instanceof Uint8Array
        ? input.byteLength
        : input instanceof ArrayBuffer
          ? input.byteLength
          : input.size;
  if (knownSize > maximumBytes) {
    throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'Interchange input byte limit exceeded');
  }
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
  if (bytes.byteLength > maximumBytes) {
    throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'Interchange input byte limit exceeded');
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
  const security = securityReport(warnings, unsupportedFeatures);
  return Object.freeze({
    format,
    document,
    diagnostics: Object.freeze(
      security.unsupportedFeatures.map((feature) =>
        Object.freeze({
          code: 'UNSUPPORTED_INTERCHANGE_FEATURE',
          severity: 'warning' as const,
          domain: 'interchange' as const,
          stage: 'decode' as const,
          message: `The imported workbook uses an unsupported feature: ${feature}`,
          details: Object.freeze({ feature }),
        }),
      ),
    ),
    security,
  });
}

export function exportResult(
  format: InterchangeFormat,
  blob: Blob,
  unsupportedFeatures: readonly string[] = [],
): WorkbookExportResult {
  return Object.freeze({
    format,
    blob,
    diagnostics: Object.freeze(
      [...new Set(unsupportedFeatures)].map((feature) =>
        Object.freeze({
          code: 'UNSUPPORTED_INTERCHANGE_FEATURE',
          severity: 'warning' as const,
          domain: 'interchange' as const,
          stage: 'serialize' as const,
          message: `The exported workbook omits an unsupported feature: ${feature}`,
          details: Object.freeze({ feature }),
        }),
      ),
    ),
  });
}
