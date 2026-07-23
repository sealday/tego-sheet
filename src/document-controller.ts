import type { WorkbookCommand } from './core/commands/workbook-command';
import {
  SpreadsheetDocumentController,
  type CommittedTransactionRecord,
  type SerializableCommandEnvelope,
  type SerializableTransactionEnvelope,
  type TransactionOptions,
  type TransactionPreview,
  type TransactionResult,
} from './core/controller/spreadsheet-document-controller';
import type { ChangeSource, WorkbookChange } from './core/types/changes';
import type { JsonValue } from './core/types/json';
import type { SpreadsheetDocument } from './document/model/document';
import type { CalculationEnvironment, FormulaFunctionRegistry } from './formula';

/**
 * Closed set of typed content commands accepted by the public transaction boundary.
 *
 * @inline
 */
export type DocumentCommand = Exclude<
  WorkbookCommand,
  | {
      /** History traversal is exposed as a controller method, not a transaction command. */
      readonly type: 'undo';
    }
  | {
      /** History traversal is exposed as a controller method, not a transaction command. */
      readonly type: 'redo';
    }
>;

/** A JSON-safe, versioned command accepted by the public document mutation boundary. */
export interface DocumentCommandEnvelope {
  /** Command envelope schema version. */
  readonly schemaVersion: 1;
  /** Caller-provided identifier unique within its transaction. */
  readonly id: string;
  /** Typed JSON-safe content command payload. */
  readonly command: DocumentCommand;
}

/** A JSON-safe atomic group of public document commands. */
export interface DocumentTransactionEnvelope {
  /** Transaction envelope schema version. */
  readonly schemaVersion: 1;
  /** Caller-provided transaction identifier. */
  readonly id: string;
  /** Revision that must still be current when the transaction starts. */
  readonly baseRevision: number;
  /** Commands applied atomically in array order. */
  readonly commands: readonly DocumentCommandEnvelope[];
  /** Optional application metadata retained with a committed transaction. */
  readonly metadata?: Readonly<Record<string, JsonValue>>;
}

/** One read-only JSON-safe operation recorded for transaction auditing. */
export type DocumentPatchOperation =
  | {
      /** Replaces the value at `path`. */
      readonly op: 'set';
      /** Object keys and array indexes leading to the replacement target. */
      readonly path: readonly (string | number)[];
      /** JSON-safe replacement value. */
      readonly value: JsonValue;
    }
  | {
      /** Deletes the object property at `path`. */
      readonly op: 'delete';
      /** Object keys leading to the property being deleted. */
      readonly path: readonly (string | number)[];
    }
  | {
      /** Removes and inserts array values at `path`. */
      readonly op: 'splice';
      /** Object keys and array indexes leading to the target array. */
      readonly path: readonly (string | number)[];
      /** Starting array index. */
      readonly index: number;
      /** Number of array items removed. */
      readonly deleteCount: number;
      /** JSON-safe array items inserted at the starting index. */
      readonly values: readonly JsonValue[];
    };

/** Machine-readable information retained while preparing a transaction. */
export interface DocumentTransactionDiagnostic {
  /** Stable diagnostic category. */
  readonly code: string;
  /** Diagnostic severity. */
  readonly severity: 'warning' | 'error';
  /** Human-readable diagnostic detail. */
  readonly message: string;
  /** Command responsible for the diagnostic when one can be identified. */
  readonly commandId?: string;
}

/** Immutable audit record returned for a committed public transaction. */
export interface DocumentCommittedTransaction extends DocumentTransactionEnvelope {
  /** Document revision created by the transaction. */
  readonly committedRevision: number;
  /** Operations that transform the prior document into the committed document. */
  readonly forwardPatches: readonly DocumentPatchOperation[];
  /** Operations that transform the committed document back to the prior document. */
  readonly inversePatches: readonly DocumentPatchOperation[];
  /** Diagnostics retained with the committed transaction. */
  readonly diagnostics: readonly DocumentTransactionDiagnostic[];
}

/** A range affected by one public document mutation. */
export interface DocumentChangedRange {
  /** Inclusive zero-based top row. */
  readonly startRow: number;
  /** Inclusive zero-based left column. */
  readonly startColumn: number;
  /** Inclusive zero-based bottom row. */
  readonly endRow: number;
  /** Inclusive zero-based right column. */
  readonly endColumn: number;
}

/** Complete mutation details for one affected worksheet. */
export interface DocumentSheetChange {
  /** Stable worksheet identifier. */
  readonly sheetId: string;
  /** Distinct mutation categories in command order. */
  readonly kinds: readonly string[];
  /** Distinct affected ranges in command order. */
  readonly ranges: readonly DocumentChangedRange[];
}

/** Aggregate public change summary for a committed mutation. */
export interface DocumentChange {
  /** Stable identifier for the mutation. */
  readonly id: string;
  /** Whether the mutation was an atomic transaction or a history operation. */
  readonly kind: 'transaction' | 'history';
  /** Interaction surface that initiated the mutation. */
  readonly source: ChangeSource;
  /** Number of commands represented by this mutation. */
  readonly commandCount: number;
  /** Every worksheet affected by the mutation. */
  readonly sheets: readonly DocumentSheetChange[];
}

/** Read-only public controller state. */
export interface DocumentControllerSnapshot {
  /** Current immutable Workbook 2.0 document. */
  readonly document: SpreadsheetDocument;
  /** Monotonic committed revision. */
  readonly revision: number;
  /** Whether an undo operation currently has an effect. */
  readonly canUndo: boolean;
  /** Whether a redo operation currently has an effect. */
  readonly canRedo: boolean;
  /** Whether new content mutations are disabled. */
  readonly readOnly: boolean;
}

/** Public event emitted after one committed mutation. */
export interface DocumentControllerEvent {
  /** Complete immutable document after the mutation. */
  readonly document: SpreadsheetDocument;
  /** Revision created by the mutation. */
  readonly revision: number;
  /** Aggregate affected-sheet and range summary. */
  readonly change: DocumentChange;
  /** Audit record when the mutation was an atomic transaction. */
  readonly transaction?: DocumentCommittedTransaction;
}

/** Context supplied synchronously to a public transaction permission gate. */
export interface DocumentTransactionPermissionContext {
  /** Isolated transaction awaiting authorization. */
  readonly transaction: DocumentTransactionEnvelope;
  /** Controller state captured before candidate execution. */
  readonly snapshot: DocumentControllerSnapshot;
}

/** Synchronous authorization hook for public document transactions. */
export type DocumentTransactionPermissionGate = (
  context: DocumentTransactionPermissionContext,
) => boolean;

/** Options shared by public transaction and command execution. */
export interface DocumentTransactionOptions {
  /** Interaction surface attributed to the mutation. */
  readonly source?: ChangeSource;
  /** Optional synchronous authorization hook called before candidate execution. */
  readonly permissionGate?: DocumentTransactionPermissionGate;
}

/** Options for executing one public command. */
export interface DocumentExecuteOptions extends DocumentTransactionOptions {
  /** Revision that must still be current, defaulting to the controller revision. */
  readonly baseRevision?: number;
}

/** Result returned by public transaction and history mutation methods. */
export interface DocumentTransactionResult {
  /** Mutation outcome. */
  readonly status: 'committed' | 'noop' | 'rejected';
  /** Stable rejection category when the mutation was rejected. */
  readonly code?: string;
  /** Human-readable rejection detail when the mutation was rejected. */
  readonly message?: string;
  /** Current or committed revision when the mutation was accepted. */
  readonly revision?: number;
  /** Complete immutable document after a committed mutation. */
  readonly document?: SpreadsheetDocument;
  /** Aggregate change summary after a committed mutation. */
  readonly change?: DocumentChange;
  /** Isolated input or committed audit record when a transaction was accepted. */
  readonly transaction?: DocumentTransactionEnvelope | DocumentCommittedTransaction;
  /** Observer failure detail when notification failed after a successful commit. */
  readonly notificationError?: string;
}

/** Side-effect-free validation result for a public transaction candidate. */
export interface DocumentTransactionPreview {
  /** Candidate outcome. */
  readonly status: 'ready' | 'noop' | 'rejected';
  /** Stable rejection category when candidate preparation failed. */
  readonly code?: string;
  /** Human-readable rejection detail when candidate preparation failed. */
  readonly message?: string;
  /** Revision against which the candidate was prepared. */
  readonly baseRevision?: number;
  /** Candidate document, without changing controller state or history. */
  readonly document?: SpreadsheetDocument;
  /** Isolated transaction that was prepared. */
  readonly transaction?: DocumentTransactionEnvelope;
}

/** Construction options for a public document controller. */
export interface DocumentControllerOptions {
  /** Whether the controller starts with content mutations disabled. */
  readonly readOnly?: boolean;
  /** Initial projected row count used by the editing engine. */
  readonly initialRowCount?: number;
  /** Initial projected column count used by the editing engine. */
  readonly initialColumnCount?: number;
  /** Deterministic formula inputs and optional F5-bridged function registry. */
  readonly calculation?: {
    /** Locale override; the document locale hint remains the fallback. */
    readonly locale?: string;
    /** IANA time zone used by date functions. */
    readonly timeZone?: string;
    /** Explicit clock sampled once per recalculation. */
    readonly clock?: CalculationEnvironment['clock'];
    /** Formula registry, including functions copied from the F5 kernel. */
    readonly functions?: FormulaFunctionRegistry;
  };
}

/** Stable public command and transaction facade for one Workbook 2.0 document. */
export interface DocumentController {
  /** Returns the current immutable public state. */
  getSnapshot(): DocumentControllerSnapshot;
  /** Executes one JSON-safe command through the atomic transaction boundary. */
  execute(
    command: DocumentCommandEnvelope,
    options?: DocumentExecuteOptions,
  ): DocumentTransactionResult;
  /** Applies a JSON-safe command group atomically. */
  transact(
    transaction: DocumentTransactionEnvelope,
    options?: DocumentTransactionOptions,
  ): DocumentTransactionResult;
  /** Prepares a transaction candidate without changing state, history, or subscriptions. */
  dryRun(
    transaction: DocumentTransactionEnvelope,
    options?: DocumentTransactionOptions,
  ): DocumentTransactionPreview;
  /** Undoes one committed command group. */
  undo(source?: ChangeSource): DocumentTransactionResult;
  /** Redoes one previously undone command group. */
  redo(source?: ChangeSource): DocumentTransactionResult;
  /** Changes whether content mutation commands are accepted. */
  setReadOnly(readOnly: boolean): void;
  /** Subscribes to committed public document changes. */
  subscribe(subscriber: (event: DocumentControllerEvent) => void): () => void;
  /** Releases subscriptions and rejects subsequent controller operations. */
  dispose(): void;
}

function toChange(change: WorkbookChange): DocumentChange {
  const details = change.aggregate?.sheets.map((entry) => ({
    sheetId: entry.sheet,
    kinds: entry.kinds,
    ranges: entry.ranges.map((range) => ({
      startRow: range.start.row,
      startColumn: range.start.column,
      endRow: range.end.row,
      endColumn: range.end.column,
    })),
  })) ?? [
    {
      sheetId: change.sheet,
      kinds: change.kind === 'history' ? ['history'] : [change.kind],
      ranges:
        change.range === undefined
          ? []
          : [
              {
                startRow: change.range.start.row,
                startColumn: change.range.start.column,
                endRow: change.range.end.row,
                endColumn: change.range.end.column,
              },
            ],
    },
  ];
  return Object.freeze({
    id: change.id,
    kind: change.kind === 'history' ? 'history' : 'transaction',
    source: change.source,
    commandCount: change.aggregate?.commandCount ?? 1,
    sheets: Object.freeze(
      details.map((entry) =>
        Object.freeze({
          ...entry,
          kinds: Object.freeze([...entry.kinds]),
          ranges: Object.freeze(entry.ranges.map((range) => Object.freeze(range))),
        }),
      ),
    ),
  });
}

function toSnapshot(controller: SpreadsheetDocumentController): DocumentControllerSnapshot {
  const snapshot = controller.getSnapshot();
  return Object.freeze({
    document: snapshot.document,
    revision: snapshot.revision,
    canUndo: snapshot.canUndo,
    canRedo: snapshot.canRedo,
    readOnly: snapshot.readOnly,
  });
}

function toOptions(
  controller: SpreadsheetDocumentController,
  options: DocumentTransactionOptions,
): TransactionOptions {
  return {
    source: options.source,
    permissionGate:
      options.permissionGate === undefined
        ? undefined
        : (context) =>
            options.permissionGate?.({
              transaction: context.transaction as unknown as DocumentTransactionEnvelope,
              snapshot: toSnapshot(controller),
            }) ?? false,
  };
}

function toRecord(record: CommittedTransactionRecord): DocumentCommittedTransaction {
  return record as unknown as DocumentCommittedTransaction;
}

function toResult(result: TransactionResult): DocumentTransactionResult {
  if (result.status === 'rejected') return result;
  if (result.status === 'noop') {
    return {
      status: 'noop',
      transaction: result.transaction as unknown as DocumentTransactionEnvelope,
      revision: result.revision,
    };
  }
  return {
    status: 'committed',
    transaction: toRecord(result.transaction),
    revision: result.revision,
    change: toChange(result.change),
    document: result.document,
    ...(result.notificationError === undefined
      ? {}
      : { notificationError: result.notificationError }),
  };
}

function toPreview(result: TransactionPreview): DocumentTransactionPreview {
  if (result.status === 'rejected') return result;
  return {
    status: result.status,
    transaction: result.transaction as unknown as DocumentTransactionEnvelope,
    baseRevision: result.baseRevision,
    document: result.document,
  };
}

class PublicDocumentController implements DocumentController {
  readonly #controller: SpreadsheetDocumentController;
  #notificationError: string | undefined;

  constructor(document: SpreadsheetDocument, options: DocumentControllerOptions) {
    this.#controller = new SpreadsheetDocumentController(document, options);
  }

  getSnapshot(): DocumentControllerSnapshot {
    return toSnapshot(this.#controller);
  }

  execute(
    command: DocumentCommandEnvelope,
    options: DocumentExecuteOptions = {},
  ): DocumentTransactionResult {
    this.#notificationError = undefined;
    return this.#withNotificationError(
      toResult(
        this.#controller.execute(command as unknown as SerializableCommandEnvelope, {
          ...toOptions(this.#controller, options),
          baseRevision: options.baseRevision,
        }),
      ),
    );
  }

  transact(
    transaction: DocumentTransactionEnvelope,
    options: DocumentTransactionOptions = {},
  ): DocumentTransactionResult {
    this.#notificationError = undefined;
    return this.#withNotificationError(
      toResult(
        this.#controller.transact(
          transaction as unknown as SerializableTransactionEnvelope,
          toOptions(this.#controller, options),
        ),
      ),
    );
  }

  dryRun(
    transaction: DocumentTransactionEnvelope,
    options: DocumentTransactionOptions = {},
  ): DocumentTransactionPreview {
    return toPreview(
      this.#controller.dryRun(
        transaction as unknown as SerializableTransactionEnvelope,
        toOptions(this.#controller, options),
      ),
    );
  }

  undo(source: ChangeSource = 'ref'): DocumentTransactionResult {
    this.#notificationError = undefined;
    const result = this.#controller.undo(source);
    if (result.status === 'noop') return { status: 'noop', revision: this.getSnapshot().revision };
    return this.#withNotificationError({
      status: 'committed',
      revision: this.getSnapshot().revision,
      document: result.commit.document,
      change: toChange(result.commit.change),
      ...(result.commit.notificationError === undefined
        ? {}
        : { notificationError: result.commit.notificationError }),
    });
  }

  redo(source: ChangeSource = 'ref'): DocumentTransactionResult {
    this.#notificationError = undefined;
    const result = this.#controller.redo(source);
    if (result.status === 'noop') return { status: 'noop', revision: this.getSnapshot().revision };
    return this.#withNotificationError({
      status: 'committed',
      revision: this.getSnapshot().revision,
      document: result.commit.document,
      change: toChange(result.commit.change),
      ...(result.commit.notificationError === undefined
        ? {}
        : { notificationError: result.commit.notificationError }),
    });
  }

  setReadOnly(readOnly: boolean): void {
    this.#controller.setReadOnly(readOnly);
  }

  subscribe(subscriber: (event: DocumentControllerEvent) => void): () => void {
    return this.#controller.subscribe((event) => {
      try {
        subscriber(
          Object.freeze({
            document: event.snapshot.document,
            revision: event.snapshot.revision,
            change: toChange(event.commit.change),
            ...(event.commit.transaction === undefined
              ? {}
              : { transaction: toRecord(event.commit.transaction) }),
          }),
        );
      } catch (error) {
        this.#notificationError =
          error instanceof Error ? error.message : 'Document controller observer failed';
      }
    });
  }

  dispose(): void {
    this.#controller.dispose();
  }

  #withNotificationError(result: DocumentTransactionResult): DocumentTransactionResult {
    if (result.status !== 'committed' || this.#notificationError === undefined) return result;
    return Object.freeze({ ...result, notificationError: this.#notificationError });
  }
}

/**
 * Creates a stable public command controller without exposing engine projections or checkpoints.
 *
 * @param document - Valid immutable Workbook 2.0 document used as initial state.
 * @param options - Optional read-only and projection sizing defaults.
 * @returns A public document command and transaction facade.
 */
export function createDocumentController(
  document: SpreadsheetDocument,
  options: DocumentControllerOptions = {},
): DocumentController {
  return new PublicDocumentController(document, options);
}
