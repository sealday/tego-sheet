import type { JsonValue } from '../core/types/json';
import type { BindingId, DocumentSheetId, ObjectId, ResourceId } from './model/ids';
import type { DocumentCellAddress, DocumentCellRange } from './model/document';

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

export interface DiagnosticLocation {
  readonly sheetId?: DocumentSheetId;
  readonly range?: DocumentCellRange;
  readonly cell?: DocumentCellAddress;
  readonly bindingId?: BindingId;
  readonly resourceId?: ResourceId;
  readonly objectId?: ObjectId;
  readonly adapterId?: string;
  readonly commandId?: string;
}

export interface Diagnostic {
  readonly code: string;
  readonly severity: 'info' | 'warning' | 'error';
  readonly domain: DiagnosticDomain;
  readonly stage: DiagnosticStage;
  readonly message: string;
  readonly location?: DiagnosticLocation;
  readonly details?: JsonValue;
  readonly cause?: unknown;
}

export type DocumentDiagnosticCode =
  | 'DOCUMENT_SCHEMA_INVALID'
  | 'UNSUPPORTED_SCHEMA_VERSION'
  | 'DUPLICATE_ID'
  | 'DANGLING_REFERENCE'
  | 'INVALID_RANGE'
  | 'INVALID_MERGE'
  | 'INVALID_EXTENSION_DATA'
  | 'DOCUMENT_LIMIT_EXCEEDED';

export interface DocumentDiagnostic extends Diagnostic {
  readonly code: DocumentDiagnosticCode;
}
