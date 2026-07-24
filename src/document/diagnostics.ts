import type { JsonValue } from '../core/types/json';
import type { BindingId, DocumentSheetId, ObjectId, ResourceId } from './model/ids';
import type { DocumentCellAddress, DocumentCellRange } from './model/document';

/** Functional domain that produced a diagnostic. */
export type DiagnosticDomain =
  | 'document'
  | 'command'
  | 'formula'
  | 'format'
  | 'validation'
  | 'view'
  | 'data'
  | 'interchange'
  | 'template'
  | 'resource'
  | 'layout'
  | 'output'
  | 'extension'
  | 'analysis'
  | 'persistence'
  | 'collaboration'
  | 'permission'
  | 'comments'
  | 'history'
  | 'ai';

/** Processing stage at which a diagnostic was produced. */
export type DiagnosticStage =
  | 'decode'
  | 'validate'
  | 'plan'
  | 'commit'
  | 'compile'
  | 'resolve'
  | 'expand'
  | 'recalculate'
  | 'layout'
  | 'render'
  | 'serialize'
  | 'load'
  | 'save'
  | 'refresh'
  | 'execute'
  | 'synchronize'
  | 'authorize'
  | 'persist'
  | 'migrate'
  | 'dispose';

/** Optional structured location associated with a diagnostic. */
export interface DiagnosticLocation {
  /** Sheet associated with the diagnostic. */
  readonly sheetId?: DocumentSheetId;
  /** Range associated with the diagnostic. */
  readonly range?: DocumentCellRange;
  /** Cell associated with the diagnostic. */
  readonly cell?: DocumentCellAddress;
  /** Binding associated with the diagnostic. */
  readonly bindingId?: BindingId;
  /** Resource associated with the diagnostic. */
  readonly resourceId?: ResourceId;
  /** Object associated with the diagnostic. */
  readonly objectId?: ObjectId;
  /** Adapter identifier associated with the diagnostic. */
  readonly adapterId?: string;
  /** Command identifier associated with the diagnostic. */
  readonly commandId?: string;
}

/** Structured machine-readable diagnostic emitted by spreadsheet operations. */
export interface Diagnostic {
  /** Stable machine-readable diagnostic code. */
  readonly code: string;
  /** Diagnostic severity. */
  readonly severity: 'info' | 'warning' | 'error';
  /** Functional domain that emitted the diagnostic. */
  readonly domain: DiagnosticDomain;
  /** Processing stage that emitted the diagnostic. */
  readonly stage: DiagnosticStage;
  /** Human-readable diagnostic message. */
  readonly message: string;
  /** Optional structured document location. */
  readonly location?: DiagnosticLocation;
  /** Optional JSON-compatible structured details. */
  readonly details?: JsonValue;
  /** Optional underlying cause for programmatic inspection. */
  readonly cause?: unknown;
}

/** Stable validation codes emitted by the Workbook 2.0 parser. */
export type DocumentDiagnosticCode =
  | 'DOCUMENT_SCHEMA_INVALID'
  | 'UNSUPPORTED_SCHEMA_VERSION'
  | 'DUPLICATE_ID'
  | 'DANGLING_REFERENCE'
  | 'INVALID_RANGE'
  | 'INVALID_MERGE'
  | 'OBJECT_ANCHOR_INVALID'
  | 'GROUP_LIMIT_EXCEEDED'
  | 'TABLE_LIMIT_EXCEEDED'
  | 'INVALID_TABLE_HEADER'
  | 'TABLE_TEMPLATE_BOUNDARY_CONFLICT'
  | 'INVALID_EXTENSION_DATA'
  | 'DOCUMENT_LIMIT_EXCEEDED';

/** Diagnostic emitted while decoding or validating a Workbook 2.0 document. */
export interface DocumentDiagnostic extends Diagnostic {
  /** Stable Workbook 2.0 diagnostic code. */
  readonly code: DocumentDiagnosticCode;
}
