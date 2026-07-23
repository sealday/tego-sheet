import { parseSpreadsheetDocument } from '../../document/parse-document';
import { projectDocumentToLegacy, projectLegacyToDocument } from './runtime-projection';
import type { SpreadsheetDocument } from '../../document/model/document';
import type { CommandResult, WorkbookCommand } from '../commands/workbook-command';
import type { ChangeSource } from '../types/changes';
import type { WorkbookChange } from '../types/changes';
import type { JsonObject } from '../types/json';
import { sheetId, type CellAddress, type SheetId } from '../types/coordinates';
import type { ValidationResult } from '../types/validation';
import { TegoSheetException } from '../errors/tego-sheet-exception';
import {
  WorkbookController,
  type ControllerSnapshot,
  type DispatchOptions,
  type WorkbookControllerOptions,
} from './workbook-controller';
import { SubscriptionStore } from './subscription-store';
import { History, type HistoryCheckpoint } from './history';
import {
  prepareSchemaCommand,
  prepareSchemaProjectionCommit,
} from '../commands/schema-command-plan';

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

export interface SpreadsheetControllerCommit<
  Result = void,
  Command extends WorkbookCommand = WorkbookCommand,
> {
  readonly command: Command;
  readonly change: import('../types/changes').WorkbookChange;
  readonly result: Result;
  readonly ['document']: SpreadsheetDocument;
  readonly transaction?: SerializableTransactionEnvelope;
}

export interface SpreadsheetControllerSnapshot extends Omit<ControllerSnapshot, 'value'> {
  readonly ['document']: SpreadsheetDocument;
  /** Read-only projection consumed only by the current engine boundary. */
  readonly projection: ControllerSnapshot['value'];
}

export interface SpreadsheetControllerEvent {
  readonly snapshot: SpreadsheetControllerSnapshot;
  readonly commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>;
}

export interface SpreadsheetDispatchOptions extends Omit<DispatchOptions, 'beforeNotify'> {
  readonly beforeNotify?: (commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>) => void;
}

export interface SpreadsheetControllerCheckpoint {
  readonly legacy: ReturnType<WorkbookController['checkpoint']>;
  readonly documentHistory: HistoryCheckpoint<SpreadsheetDocument, null>;
  readonly ['document']: SpreadsheetDocument;
}

/** A JSON-serializable, versioned document command submitted to the transaction boundary. */
export interface SerializableCommandEnvelope {
  readonly schemaVersion: 1;
  readonly id: string;
  readonly command: WorkbookCommand;
}

/** A JSON-serializable atomic group of document commands. */
export interface SerializableTransactionEnvelope {
  readonly schemaVersion: 1;
  readonly id: string;
  readonly baseRevision: number;
  readonly commands: readonly SerializableCommandEnvelope[];
  readonly metadata?: JsonObject;
}

/** Context supplied to a transaction permission gate before candidate execution starts. */
export interface TransactionPermissionContext {
  readonly transaction: SerializableTransactionEnvelope;
  readonly snapshot: SpreadsheetControllerSnapshot;
}

/** Synchronous authorization hook for document transactions. */
export type TransactionPermissionGate = (context: TransactionPermissionContext) => boolean;

export interface TransactionOptions {
  readonly source?: ChangeSource;
  readonly permissionGate?: TransactionPermissionGate;
}

export interface ExecuteOptions extends TransactionOptions {
  readonly baseRevision?: number;
}

export type TransactionRejectionCode =
  | 'COMMAND_SCHEMA_INVALID'
  | 'COMMAND_NOT_ALLOWED'
  | 'REVISION_CONFLICT'
  | 'TRANSACTION_INVARIANT_FAILED'
  | 'TRANSACTION_LIMIT_EXCEEDED';

export interface TransactionRejection {
  readonly status: 'rejected';
  readonly code: TransactionRejectionCode;
  readonly message: string;
}

export interface TransactionCommit {
  readonly status: 'committed';
  readonly transaction: SerializableTransactionEnvelope;
  readonly revision: number;
  readonly change: WorkbookChange;
  readonly ['document']: SpreadsheetDocument;
}

export interface TransactionNoop {
  readonly status: 'noop';
  readonly transaction: SerializableTransactionEnvelope;
  readonly revision: number;
}

export type TransactionResult = TransactionCommit | TransactionNoop | TransactionRejection;

export type TransactionPreview =
  | TransactionRejection
  | {
      readonly status: 'ready' | 'noop';
      readonly transaction: SerializableTransactionEnvelope;
      readonly baseRevision: number;
      readonly ['document']: SpreadsheetDocument;
    };

const MAX_TRANSACTION_COMMANDS = 1_000;

function captureJsonValue(value: unknown, seen = new Set<object>()): unknown {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'object' || seen.has(value)) {
    throw new TypeError('Value is not finite acyclic JSON');
  }
  seen.add(value);
  try {
    if (Array.isArray(value)) {
      const descriptors = Object.getOwnPropertyDescriptors(value);
      const output: unknown[] = [];
      for (let index = 0; index < value.length; index += 1) {
        const descriptor = descriptors[String(index)];
        if (
          descriptor === undefined ||
          !descriptor.enumerable ||
          !Object.hasOwn(descriptor, 'value')
        ) {
          throw new TypeError('JSON arrays must contain data properties');
        }
        output.push(captureJsonValue(descriptor.value, seen));
      }
      return Object.freeze(output);
    }
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new TypeError('JSON objects must use a plain prototype');
    }
    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw new TypeError('JSON objects cannot contain symbol keys');
    }
    const output = Object.create(null) as Record<string, unknown>;
    for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
      if (!descriptor.enumerable) continue;
      if (!Object.hasOwn(descriptor, 'value')) {
        throw new TypeError('JSON objects cannot contain accessors');
      }
      Object.defineProperty(output, key, {
        configurable: false,
        enumerable: true,
        value: captureJsonValue(descriptor.value, seen),
        writable: false,
      });
    }
    return Object.freeze(output);
  } finally {
    seen.delete(value);
  }
}

function snapshotTransaction(
  value: SerializableTransactionEnvelope,
): SerializableTransactionEnvelope | TransactionRejection {
  let transaction: SerializableTransactionEnvelope;
  try {
    const captured = captureJsonValue(value);
    if (captured === null || typeof captured !== 'object' || Array.isArray(captured)) {
      throw new TypeError('Transaction must be an object');
    }
    transaction = captured as SerializableTransactionEnvelope;
  } catch {
    return {
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
      message: 'Transaction could not be isolated',
    };
  }
  if (
    transaction.schemaVersion !== 1 ||
    typeof transaction.id !== 'string' ||
    transaction.id.length === 0 ||
    !Number.isSafeInteger(transaction.baseRevision) ||
    transaction.baseRevision < 0 ||
    !Array.isArray(transaction.commands)
  ) {
    return {
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
      message: 'Transaction envelope is invalid',
    };
  }
  if (transaction.commands.length > MAX_TRANSACTION_COMMANDS) {
    return {
      status: 'rejected',
      code: 'TRANSACTION_LIMIT_EXCEEDED',
      message: `Transaction exceeds ${MAX_TRANSACTION_COMMANDS} commands`,
    };
  }
  if (
    transaction.commands.some(
      (entry) =>
        entry.schemaVersion !== 1 ||
        typeof entry.id !== 'string' ||
        entry.id.length === 0 ||
        entry.command === null ||
        typeof entry.command !== 'object' ||
        entry.command.type === 'undo' ||
        entry.command.type === 'redo',
    )
  ) {
    return {
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
      message: 'Transaction command envelope is invalid',
    };
  }
  if (new Set(transaction.commands.map((entry) => entry.id)).size !== transaction.commands.length) {
    return {
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
      message: 'Transaction command IDs must be unique',
    };
  }
  return transaction;
}

/** @internal */
export function cloneFrozenDocumentValue<T>(value: T): T {
  if (value === null || typeof value !== 'object') return value;
  const output = (Array.isArray(value) ? [] : Object.create(null)) as Record<string, unknown>;
  for (const [key, item] of Object.entries(value)) {
    Object.defineProperty(output, key, {
      configurable: false,
      enumerable: true,
      value: cloneFrozenDocumentValue(item),
      writable: false,
    });
  }
  return Object.freeze(output) as T;
}

/**
 * Owns the single schema 2 runtime truth while adapting existing operations at one private boundary.
 */
export class SpreadsheetDocumentController {
  private currentDocument: SpreadsheetDocument;
  private checkpointDocument: SpreadsheetDocument | undefined;
  private readonly documentHistory = new History<SpreadsheetDocument, null>();
  private readonly legacy: WorkbookController;
  private readonly subscriptions = new SubscriptionStore<SpreadsheetControllerEvent>();

  constructor(input: SpreadsheetDocument, options: WorkbookControllerOptions = {}) {
    const parsed = parseSpreadsheetDocument(input);
    if (!parsed.ok) {
      throw new TegoSheetException({
        code: 'INVALID_DATA',
        message: 'Spreadsheet document is invalid',
        recoverable: false,
        cause: parsed.diagnostics,
      });
    }
    const { ['document']: parsedDocument } = parsed;
    this.currentDocument = parsedDocument;
    this.legacy = new WorkbookController(projectDocumentToLegacy(this.currentDocument), {
      ...options,
      sheetIds: this.currentDocument.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    });
  }

  get historySize() {
    return this.documentHistory.size;
  }

  get canUndo(): boolean {
    return this.documentHistory.canUndo;
  }

  get canRedo(): boolean {
    return this.documentHistory.canRedo;
  }

  getDocument(): SpreadsheetDocument {
    return cloneFrozenDocumentValue(this.currentDocument);
  }

  getSheetIds(): readonly SheetId[] {
    return this.legacy.getSheetIds();
  }

  getInitializationDefaults() {
    return this.legacy.getInitializationDefaults();
  }

  getCellText(address: CellAddress): string {
    return this.legacy.getCellText(address);
  }

  getSnapshot(): SpreadsheetControllerSnapshot {
    const snapshot = this.legacy.getSnapshot();
    const { value: projection, ...metadata } = snapshot;
    return cloneFrozenDocumentValue({
      ...metadata,
      ['document']: this.currentDocument,
      projection,
    });
  }

  validate(): ValidationResult {
    return this.legacy.validate();
  }

  subscribe(subscriber: (event: SpreadsheetControllerEvent) => void): () => void {
    return this.subscriptions.subscribe(subscriber);
  }

  execute(command: SerializableCommandEnvelope, options: ExecuteOptions = {}): TransactionResult {
    return this.transact(
      {
        schemaVersion: 1,
        id: `transaction:${command.id}`,
        baseRevision: options.baseRevision ?? this.getSnapshot().revision,
        commands: [command],
      },
      options,
    );
  }

  transact(
    input: SerializableTransactionEnvelope,
    options: TransactionOptions = {},
  ): TransactionResult {
    const checked = this.checkTransaction(input, options);
    if ('status' in checked) return checked;
    const transaction = checked;
    const checkpoint = this.checkpoint();
    const source = options.source ?? 'ref';
    let lastCommit: SpreadsheetControllerCommit<unknown, WorkbookCommand> | undefined;
    let event: SpreadsheetControllerEvent;
    let result: TransactionCommit;
    try {
      for (const envelope of transaction.commands) {
        const outcome = this.dispatch(envelope.command, source, { notify: false });
        if (outcome.status === 'committed') lastCommit = outcome.commit;
      }
      if (lastCommit === undefined) {
        this.restore(checkpoint);
        return {
          status: 'noop',
          transaction,
          revision: checkpoint.legacy.revision,
        };
      }
      const candidate = this.currentDocument;
      this.documentHistory.restore(checkpoint.documentHistory);
      this.documentHistory.record({
        before: checkpoint.document,
        after: candidate,
        metadata: null,
      });
      const finalized = this.legacy.finalizeTransaction(
        checkpoint.legacy,
        lastCommit.command,
        source,
        {
          kind: 'transaction',
          sheet: lastCommit.change.sheet,
        },
      );
      if (finalized.status === 'noop') {
        this.restore(checkpoint);
        return {
          status: 'noop',
          transaction,
          revision: checkpoint.legacy.revision,
        };
      }
      const commit = cloneFrozenDocumentValue({
        ...lastCommit,
        change: finalized.commit.change,
        ['document']: candidate,
        transaction,
      }) as SpreadsheetControllerCommit<unknown, WorkbookCommand>;
      this.currentDocument = candidate;
      event = cloneFrozenDocumentValue({
        snapshot: this.getSnapshot(),
        commit,
      }) as SpreadsheetControllerEvent;
      result = cloneFrozenDocumentValue({
        status: 'committed',
        transaction,
        revision: this.getSnapshot().revision,
        change: finalized.commit.change,
        ['document']: candidate,
      }) as TransactionCommit;
    } catch (error) {
      this.restore(checkpoint);
      return this.rejectTransaction(error);
    }
    this.subscriptions.publish(event);
    return result;
  }

  dryRun(
    input: SerializableTransactionEnvelope,
    options: TransactionOptions = {},
  ): TransactionPreview {
    const checked = this.checkTransaction(input, options);
    if ('status' in checked) return checked;
    const transaction = checked;
    const checkpoint = this.checkpoint();
    const source = options.source ?? 'ref';
    try {
      let changed = false;
      for (const envelope of transaction.commands) {
        const outcome = this.dispatch(envelope.command, source, { notify: false });
        if (outcome.status === 'committed') changed = true;
      }
      const candidate = this.getDocument();
      return cloneFrozenDocumentValue({
        status: changed ? 'ready' : 'noop',
        transaction,
        baseRevision: transaction.baseRevision,
        ['document']: candidate,
      }) as TransactionPreview;
    } catch (error) {
      return this.rejectTransaction(error);
    } finally {
      this.restore(checkpoint);
    }
  }

  dispatch<Command extends WorkbookCommand>(
    command: Command,
    source: ChangeSource,
    options: SpreadsheetDispatchOptions = {},
  ):
    | { readonly status: 'noop' }
    | {
        readonly status: 'committed';
        readonly commit: SpreadsheetControllerCommit<CommandResult<Command>, Command>;
      } {
    this.legacy.assertCommand(command);
    const plan = prepareSchemaCommand(this.currentDocument, command, this.legacy.getSheetIds());
    const plannedProjection = projectDocumentToLegacy(plan.document);
    const historyCheckpoint = this.documentHistory.checkpoint();
    let preparedDocument: SpreadsheetDocument | undefined;
    let preparedCommit: SpreadsheetControllerCommit<CommandResult<Command>, Command> | undefined;
    let outcome: ReturnType<WorkbookController['dispatch']>;
    const beforeNotify: NonNullable<DispatchOptions['beforeNotify']> = (legacyCommit) => {
      let candidate: SpreadsheetDocument;
      if (command.type === 'undo') {
        const entry = this.documentHistory.undo();
        if (entry === null) throw new Error('Schema history is not aligned for undo');
        candidate = entry.before;
      } else if (command.type === 'redo') {
        const entry = this.documentHistory.redo();
        if (entry === null) throw new Error('Schema history is not aligned for redo');
        candidate = entry.after;
      } else {
        candidate = projectLegacyToDocument(
          plannedProjection,
          legacyCommit.value,
          plan.document,
          this.legacy.getSheetIds(),
          plan.authoritativeInputs,
          plan.authoritativeValidations,
        );
        this.documentHistory.record({
          before: this.currentDocument,
          after: candidate,
          metadata: null,
        });
      }
      this.legacy.reconcileProjection(
        projectDocumentToLegacy(candidate),
        this.legacy.getSheetIds(),
        command.type === 'undo' ? 'undo' : command.type === 'redo' ? 'redo' : 'commit',
      );
      const commit = cloneFrozenDocumentValue({
        command: legacyCommit.command,
        change: legacyCommit.change,
        result: legacyCommit.result,
        ['document']: candidate,
      }) as SpreadsheetControllerCommit<CommandResult<Command>, Command>;
      this.checkpointDocument = candidate;
      try {
        options.beforeNotify?.(commit as never);
        preparedDocument = candidate;
        preparedCommit = commit;
      } finally {
        this.checkpointDocument = undefined;
      }
    };
    const legacyOptions: DispatchOptions = {
      ...options,
      beforeNotify,
      notify: false,
    };
    try {
      outcome = this.legacy.dispatch(command, source, legacyOptions);
      if (outcome.status === 'noop' && !sameJson(plan.document, this.currentDocument)) {
        if (command.type !== 'paste-internal' && command.type !== 'autofill') {
          throw new Error(`Schema-only commit is not supported for ${command.type}`);
        }
        const projectionCommit = prepareSchemaProjectionCommit(
          command,
          this.legacy.getValue(),
          this.legacy.getSheetIds(),
          options.capturePasteValues !== false,
        );
        outcome = this.legacy.commitProjection(
          command,
          source,
          {
            value: plannedProjection,
            sheetIds: this.legacy.getSheetIds(),
            result: projectionCommit.result as CommandResult<Command>,
            kind: projectionCommit.kind,
            sheet: projectionCommit.sheet,
            range: projectionCommit.range,
          },
          legacyOptions,
        );
      }
    } catch (error) {
      this.documentHistory.restore(historyCheckpoint);
      throw error;
    }
    if (outcome.status === 'noop') return outcome;
    if (preparedDocument === undefined || preparedCommit === undefined) {
      throw new Error('Spreadsheet document transaction was not prepared');
    }
    this.currentDocument = preparedDocument;
    const commit = preparedCommit;
    if (options.notify !== false) {
      const event = cloneFrozenDocumentValue({
        snapshot: this.getSnapshot(),
        commit,
      }) as SpreadsheetControllerEvent;
      this.subscriptions.publish(event);
    }
    return { status: 'committed', commit };
  }

  undo(source: ChangeSource = 'ref', options: SpreadsheetDispatchOptions = {}) {
    return this.dispatch({ type: 'undo' }, source, options);
  }

  redo(source: ChangeSource = 'ref', options: SpreadsheetDispatchOptions = {}) {
    return this.dispatch({ type: 'redo' }, source, options);
  }

  checkpoint(): SpreadsheetControllerCheckpoint {
    return {
      legacy: this.legacy.checkpoint(),
      documentHistory: this.documentHistory.checkpoint(),
      ['document']: this.checkpointDocument ?? this.currentDocument,
    };
  }

  restore(checkpoint: SpreadsheetControllerCheckpoint): void {
    this.legacy.restore(checkpoint.legacy);
    this.documentHistory.restore(checkpoint.documentHistory);
    const { ['document']: checkpointDocument } = checkpoint;
    this.currentDocument = checkpointDocument;
  }

  replace(input: SpreadsheetDocument): void {
    const parsed = parseSpreadsheetDocument(input);
    if (!parsed.ok) {
      throw new TegoSheetException({
        code: 'INVALID_DATA',
        message: 'Spreadsheet document is invalid',
        recoverable: true,
        cause: parsed.diagnostics,
      });
    }
    const { ['document']: parsedDocument } = parsed;
    const projection = projectDocumentToLegacy(parsedDocument);
    this.legacy.replace(
      projection,
      parsedDocument.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    );
    this.currentDocument = parsedDocument;
    this.documentHistory.clear();
  }

  setReadOnly(readOnly: boolean): void {
    this.legacy.setReadOnly(readOnly);
  }

  dispose(): void {
    this.subscriptions.dispose();
    this.legacy.dispose();
  }

  private checkTransaction(
    input: SerializableTransactionEnvelope,
    options: TransactionOptions,
  ): SerializableTransactionEnvelope | TransactionRejection {
    const transaction = snapshotTransaction(input);
    if ('status' in transaction) return transaction;
    if (
      options.source !== undefined &&
      ![
        'keyboard',
        'pointer',
        'touch',
        'toolbar',
        'sheet-tabs',
        'context-menu',
        'clipboard',
        'ref',
      ].includes(options.source)
    ) {
      return {
        status: 'rejected',
        code: 'COMMAND_SCHEMA_INVALID',
        message: 'Transaction source is invalid',
      };
    }
    if (transaction.baseRevision !== this.getSnapshot().revision) {
      return {
        status: 'rejected',
        code: 'REVISION_CONFLICT',
        message: `Expected revision ${transaction.baseRevision}, current revision is ${this.getSnapshot().revision}`,
      };
    }
    try {
      if (
        options.permissionGate?.({
          transaction,
          snapshot: this.getSnapshot(),
        }) === false
      ) {
        return {
          status: 'rejected',
          code: 'COMMAND_NOT_ALLOWED',
          message: 'Transaction was denied by the permission gate',
        };
      }
    } catch (error) {
      return {
        status: 'rejected',
        code: 'COMMAND_NOT_ALLOWED',
        message: error instanceof Error ? error.message : 'Transaction permission gate failed',
      };
    }
    return transaction;
  }

  private rejectTransaction(error: unknown): TransactionRejection {
    return {
      status: 'rejected',
      code:
        error instanceof TegoSheetException && error.code === 'INVALID_COMMAND'
          ? 'COMMAND_SCHEMA_INVALID'
          : 'TRANSACTION_INVARIANT_FAILED',
      message: error instanceof Error ? error.message : 'Transaction failed',
    };
  }
}
