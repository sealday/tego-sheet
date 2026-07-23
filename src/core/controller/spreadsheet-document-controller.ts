import { parseSpreadsheetDocument } from '../../document/parse-document';
import { projectDocumentToLegacy, projectLegacyToDocument } from './runtime-projection';
import type { CellInput, FilterView, SpreadsheetDocument } from '../../document/model/document';
import {
  createFormulaEngine,
  createFormulaFunctionRegistry,
  type CalculationEnvironment,
  type FormulaEngine,
  type FormulaFunctionRegistry,
  type FormulaProgram,
  type FormulaValue,
} from '../../formula';
import type { CommandResult, WorkbookCommand } from '../commands/workbook-command';
import type {
  ChangeSource,
  TransactionChangeAggregate,
  TransactionSheetChange,
  WorkbookChange,
} from '../types/changes';
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
import { applyDocumentPatch, createDocumentPatch, type DocumentPatch } from './document-patch';
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
  readonly transaction?: CommittedTransactionRecord;
  readonly notificationError?: string;
}

export interface SpreadsheetControllerSnapshot extends Omit<ControllerSnapshot, 'value'> {
  readonly ['document']: SpreadsheetDocument;
  /** Read-only projection consumed only by the current engine boundary. */
  readonly projection: ControllerSnapshot['value'];
  /** JSON-safe typed formula values for read-only presentation adapters. */
  readonly calculation: {
    /** Formula calculation revision matching this snapshot. */
    readonly revision: number;
    /** Stable address/value entries without exposing mutable Map state. */
    readonly values: readonly {
      readonly address: string;
      readonly value: FormulaValue;
    }[];
  };
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
  readonly documentHistory: HistoryCheckpoint<DocumentPatch, null>;
  readonly ['document']: SpreadsheetDocument;
  readonly [spreadsheetCheckpointOwner]: object;
}

const spreadsheetCheckpointOwner: unique symbol = Symbol(
  'tego-sheet.spreadsheet-controller-checkpoint-owner',
);

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
  readonly transaction: CommittedTransactionRecord;
  readonly revision: number;
  readonly change: WorkbookChange;
  readonly ['document']: SpreadsheetDocument;
  readonly notificationError?: string;
}

/** Machine-readable information produced while preparing a committed transaction. */
export interface TransactionDiagnostic {
  /** Stable diagnostic category. */
  readonly code: string;
  /** Diagnostic severity. */
  readonly severity: 'warning' | 'error';
  /** Human-readable diagnostic detail. */
  readonly message: string;
  /** Command responsible for the diagnostic when one command can be identified. */
  readonly commandId?: string;
}

/** Immutable audit record produced for a successfully committed transaction. */
export interface CommittedTransactionRecord extends SerializableTransactionEnvelope {
  /** Document revision created by this commit. */
  readonly committedRevision: number;
  /** JSON-safe operations that transform the pre-commit document into the committed document. */
  readonly forwardPatches: DocumentPatch['operations'];
  /** JSON-safe operations that transform the committed document back to its pre-commit state. */
  readonly inversePatches: DocumentPatch['operations'];
  /** Preparation diagnostics retained with the committed audit record. */
  readonly diagnostics: readonly TransactionDiagnostic[];
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
const MAX_TRANSACTION_BYTES = 4 * 1_024 * 1_024;

function calculationEnvironment(
  document: SpreadsheetDocument,
  revision: number,
  options: SpreadsheetCalculationOptions,
): CalculationEnvironment {
  return {
    locale: options.locale ?? document.workbook.settings.localeHint ?? 'en-US',
    timeZone: options.timeZone ?? 'UTC',
    dateSystem: document.workbook.settings.dateSystem,
    clock: options.clock ?? { now: () => 0 },
    tick: revision,
    functionRegistryVersion: options.functions.version,
    resolveVolatile: options.resolveVolatile,
  };
}

/** Host-controlled deterministic inputs used by the schema 2 calculation runtime. */
export interface SpreadsheetCalculationOptions {
  /** Locale override; the document locale hint remains the fallback. */
  readonly locale?: string;
  /** IANA time zone used by date functions. */
  readonly timeZone?: string;
  /** Explicit clock sampled once per recalculation. */
  readonly clock?: CalculationEnvironment['clock'];
  /** Formula registry, including any functions bridged from the F5 kernel. */
  readonly functions: FormulaFunctionRegistry;
  /** @internal Whether both volatile inputs were explicitly supplied by the host. */
  readonly resolveVolatile: boolean;
}

export interface SpreadsheetDocumentControllerOptions extends WorkbookControllerOptions {
  readonly calculation?: Partial<SpreadsheetCalculationOptions>;
}

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
    !Array.isArray(transaction.commands) ||
    (transaction.metadata !== undefined &&
      (transaction.metadata === null ||
        typeof transaction.metadata !== 'object' ||
        Array.isArray(transaction.metadata)))
  ) {
    return {
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
      message: 'Transaction envelope is invalid',
    };
  }
  if (new TextEncoder().encode(JSON.stringify(transaction)).byteLength > MAX_TRANSACTION_BYTES) {
    return {
      status: 'rejected',
      code: 'TRANSACTION_LIMIT_EXCEEDED',
      message: `Transaction exceeds ${MAX_TRANSACTION_BYTES} encoded bytes`,
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

function aggregateTransactionChanges(
  commits: readonly SpreadsheetControllerCommit<unknown, WorkbookCommand>[],
): TransactionChangeAggregate {
  const sheets = new Map<
    SheetId,
    {
      kinds: Array<Exclude<WorkbookChange['kind'], 'transaction'>>;
      ranges: NonNullable<WorkbookChange['range']>[];
    }
  >();
  for (const commit of commits) {
    const current = sheets.get(commit.change.sheet) ?? {
      kinds: [],
      ranges: [],
    };
    if (commit.change.kind !== 'transaction' && !current.kinds.includes(commit.change.kind)) {
      current.kinds.push(commit.change.kind);
    }
    if (
      commit.change.range !== undefined &&
      !current.ranges.some((range) => sameJson(range, commit.change.range))
    ) {
      current.ranges.push(commit.change.range);
    }
    sheets.set(commit.change.sheet, current);
  }
  return {
    commandCount: commits.length,
    sheets: [...sheets].map(
      ([sheet, detail]): TransactionSheetChange => ({
        sheet,
        kinds: detail.kinds,
        ranges: detail.ranges,
      }),
    ),
  };
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
  private readonly documentHistory = new History<DocumentPatch, null>();
  private readonly formulaEngine: FormulaEngine;
  private readonly calculationOptions: SpreadsheetCalculationOptions;
  private formulaProgram: FormulaProgram;
  private formulaValues: ReadonlyMap<string, FormulaValue>;
  private readonly activeFilterViews = new Map<SheetId, string>();
  private filterViewRevision = 0;
  private readonly legacy: WorkbookController;
  private readonly subscriptions = new SubscriptionStore<SpreadsheetControllerEvent>();
  private permissionGateActive = false;
  private commitMutationActive = false;
  private checkpointOwner: object = Object.freeze({});
  private checkpoints = new WeakSet<SpreadsheetControllerCheckpoint>();

  constructor(input: SpreadsheetDocument, options: SpreadsheetDocumentControllerOptions = {}) {
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
    const functions = options.calculation?.functions ?? createFormulaFunctionRegistry();
    this.calculationOptions = {
      functions,
      locale: options.calculation?.locale,
      timeZone: options.calculation?.timeZone ?? 'UTC',
      clock: options.calculation?.clock ?? { now: () => 0 },
      resolveVolatile:
        options.calculation?.clock !== undefined && options.calculation.timeZone !== undefined,
    };
    this.formulaEngine = createFormulaEngine({ functions });
    this.formulaProgram = this.formulaEngine.compile(this.currentDocument);
    this.formulaValues = this.formulaEngine.recalculate(
      this.formulaProgram,
      [],
      calculationEnvironment(this.currentDocument, 0, this.calculationOptions),
    ).values;
    this.legacy = new WorkbookController(
      projectDocumentToLegacy(this.currentDocument, this.formulaValues),
      {
        ...options,
        sheetIds: this.currentDocument.workbook.sheets.map((sheet) => sheetId(sheet.id)),
      },
    );
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

  /** Selects one saved view as session-only state without changing document history. */
  activateFilterView(sheet: SheetId, viewId: string): void {
    const index = this.getSheetIds().findIndex((candidate) => candidate === sheet);
    const documentSheet = this.currentDocument.workbook.sheets[index];
    if (documentSheet === undefined) throw new RangeError(`Unknown sheet ID: ${sheet}`);
    if (!documentSheet.filterViews.some((view) => view.id === viewId)) {
      throw new RangeError(`Unknown filter view ID: ${viewId}`);
    }
    if (this.activeFilterViews.get(sheet) === viewId) return;
    this.activeFilterViews.set(sheet, viewId);
    this.filterViewRevision += 1;
  }

  /** Clears the selected session view for one worksheet. */
  deactivateFilterView(sheet: SheetId): void {
    if (!this.activeFilterViews.delete(sheet)) return;
    this.filterViewRevision += 1;
  }

  /** Returns the selected immutable view definition for one worksheet. */
  getActiveFilterView(sheet: SheetId): FilterView | undefined {
    const index = this.getSheetIds().findIndex((candidate) => candidate === sheet);
    const viewId = this.activeFilterViews.get(sheet);
    if (index < 0 || viewId === undefined) return undefined;
    return this.currentDocument.workbook.sheets[index]?.filterViews.find(
      (view) => view.id === viewId,
    );
  }

  /** Monotonic session revision used to invalidate derived presentation caches. */
  getFilterViewRevision(): number {
    return this.filterViewRevision;
  }

  getSnapshot(): SpreadsheetControllerSnapshot {
    const snapshot = this.legacy.getSnapshot();
    const { value: projection, ...metadata } = snapshot;
    return cloneFrozenDocumentValue({
      ...metadata,
      ['document']: this.currentDocument,
      calculation: {
        revision: snapshot.revision,
        values: [...this.formulaValues]
          .sort(([left], [right]) => left.localeCompare(right))
          .map(([address, value]) => ({ address, value })),
      },
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
        id: 'transaction:execute',
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
    const commits: SpreadsheetControllerCommit<unknown, WorkbookCommand>[] = [];
    let event: SpreadsheetControllerEvent;
    let result: TransactionCommit;
    try {
      for (const envelope of transaction.commands) {
        const outcome = this.dispatch(envelope.command, source, {
          notify: false,
        });
        if (outcome.status === 'committed') commits.push(outcome.commit);
      }
      const lastCommit = commits.at(-1);
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
      this.documentHistory.record(this.createDocumentHistoryEntry(checkpoint.document, candidate));
      const aggregate = aggregateTransactionChanges(commits);
      const finalized = this.legacy.finalizeTransaction(
        checkpoint.legacy,
        lastCommit.command,
        source,
        {
          kind: 'transaction',
          sheet: lastCommit.change.sheet,
          aggregate,
        },
        !sameJson(checkpoint.document, candidate),
      );
      if (finalized.status === 'noop') {
        this.restore(checkpoint);
        return {
          status: 'noop',
          transaction,
          revision: checkpoint.legacy.revision,
        };
      }
      const change = finalized.commit.change;
      const transactionRecord = cloneFrozenDocumentValue({
        ...transaction,
        committedRevision: this.getSnapshot().revision,
        forwardPatches: createDocumentPatch(checkpoint.document, candidate).operations,
        inversePatches: createDocumentPatch(candidate, checkpoint.document).operations,
        diagnostics: [],
      }) as CommittedTransactionRecord;
      const commit = cloneFrozenDocumentValue({
        ...lastCommit,
        change,
        ['document']: candidate,
        transaction: transactionRecord,
      }) as SpreadsheetControllerCommit<unknown, WorkbookCommand>;
      this.currentDocument = candidate;
      event = cloneFrozenDocumentValue({
        snapshot: this.getSnapshot(),
        commit,
      }) as SpreadsheetControllerEvent;
      result = cloneFrozenDocumentValue({
        status: 'committed',
        transaction: transactionRecord,
        revision: this.getSnapshot().revision,
        change,
        ['document']: candidate,
      }) as TransactionCommit;
    } catch (error) {
      this.restore(checkpoint);
      return this.rejectTransaction(error);
    }
    try {
      this.subscriptions.publish(event);
      return result;
    } catch (error) {
      return cloneFrozenDocumentValue({
        ...result,
        notificationError: error instanceof Error ? error.message : 'Transaction observer failed',
      }) as TransactionCommit;
    }
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
        const outcome = this.dispatch(envelope.command, source, {
          notify: false,
        });
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
    this.assertNoActiveCommitMutation();
    if (this.permissionGateActive) {
      throw new TegoSheetException({
        code: 'INVALID_COMMAND',
        message: 'Permission gates cannot mutate the document controller',
        recoverable: true,
      });
    }
    this.legacy.assertCommand(command);
    const rollbackCheckpoint = this.checkpoint();
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
        candidate = this.applyPatchToDocument(this.currentDocument, entry.before);
      } else if (command.type === 'redo') {
        const entry = this.documentHistory.redo();
        if (entry === null) throw new Error('Schema history is not aligned for redo');
        candidate = this.applyPatchToDocument(this.currentDocument, entry.after);
      } else {
        candidate = projectLegacyToDocument(
          plannedProjection,
          legacyCommit.value,
          plan.document,
          this.legacy.getSheetIds(),
          plan.authoritativeInputs,
          plan.authoritativeValidations,
        );
        this.documentHistory.record(
          this.createDocumentHistoryEntry(this.currentDocument, candidate),
        );
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
    this.commitMutationActive = true;
    try {
      outcome = this.legacy.dispatch(command, source, legacyOptions);
      if (outcome.status === 'noop' && !sameJson(plan.document, this.currentDocument)) {
        if (
          command.type !== 'paste-internal' &&
          command.type !== 'autofill' &&
          command.type !== 'set-filter-view' &&
          command.type !== 'remove-filter-view' &&
          command.type !== 'set-conditional-format' &&
          command.type !== 'remove-conditional-format' &&
          command.type !== 'set-sheet-object' &&
          command.type !== 'remove-sheet-object'
        ) {
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
            ...(projectionCommit.range === undefined ? {} : { range: projectionCommit.range }),
          },
          legacyOptions,
        );
      }
    } catch (error) {
      this.documentHistory.restore(historyCheckpoint);
      throw error;
    } finally {
      this.commitMutationActive = false;
    }
    if (outcome.status === 'noop') return outcome;
    if (preparedDocument === undefined || preparedCommit === undefined) {
      throw new Error('Spreadsheet document transaction was not prepared');
    }
    const commit = preparedCommit;
    try {
      this.refreshFormulaCalculation(preparedDocument, commit.change);
      this.legacy.reconcileProjection(
        projectDocumentToLegacy(preparedDocument, this.formulaValues),
        this.legacy.getSheetIds(),
        command.type === 'undo' ? 'undo' : command.type === 'redo' ? 'redo' : 'commit',
      );
      this.currentDocument = preparedDocument;
    } catch (error) {
      this.restore(rollbackCheckpoint);
      throw error;
    }
    if (options.notify !== false) {
      const event = cloneFrozenDocumentValue({
        snapshot: this.getSnapshot(),
        commit,
      }) as SpreadsheetControllerEvent;
      try {
        this.subscriptions.publish(event);
      } catch (error) {
        return {
          status: 'committed',
          commit: cloneFrozenDocumentValue({
            ...commit,
            notificationError:
              error instanceof Error ? error.message : 'Spreadsheet observer failed',
          }),
        };
      }
    }
    return { status: 'committed', commit };
  }

  private refreshFormulaCalculation(document: SpreadsheetDocument, change: WorkbookChange): void {
    const impacts =
      change.kind === 'transaction'
        ? change.aggregate?.sheets.flatMap((entry) =>
            entry.kinds.every((kind) => kind === 'cell')
              ? entry.ranges.map((range) => ({ sheet: entry.sheet, range }))
              : [],
          )
        : change.kind === 'cell' && change.range !== undefined
          ? [{ sheet: change.sheet, range: change.range }]
          : [];
    const incremental =
      impacts !== undefined &&
      impacts.length > 0 &&
      (change.kind === 'cell' ||
        change.aggregate?.sheets.every((entry) => entry.kinds.every((kind) => kind === 'cell')) ===
          true);
    if (!incremental) {
      const formulaProgram = this.formulaEngine.compile(document);
      const formulaValues = this.formulaEngine.recalculate(
        formulaProgram,
        [],
        calculationEnvironment(document, this.getSnapshot().revision, this.calculationOptions),
      ).values;
      this.formulaProgram = formulaProgram;
      this.formulaValues = formulaValues;
      return;
    }

    const sheetById = new Map(document.workbook.sheets.map((sheet) => [String(sheet.id), sheet]));
    const dependencies = impacts.flatMap(({ sheet: changedSheet, range }) => {
      const sheet = sheetById.get(String(changedSheet));
      if (sheet === undefined) return [];
      const inputs = new Map(
        sheet.cells.map(({ row, column, cell }) => [`${row}:${column}`, cell.input]),
      );
      const output: Array<{
        sheetId: string;
        row: number;
        column: number;
        input: CellInput;
      }> = [];
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let column = range.start.column; column <= range.end.column; column += 1) {
          output.push({
            sheetId: sheet.id,
            row,
            column,
            input: inputs.get(`${row}:${column}`) ?? { type: 'blank' },
          });
        }
      }
      return output;
    });
    this.formulaValues = this.formulaEngine.recalculate(
      this.formulaProgram,
      dependencies,
      calculationEnvironment(document, this.getSnapshot().revision, this.calculationOptions),
    ).values;
  }

  undo(source: ChangeSource = 'ref', options: SpreadsheetDispatchOptions = {}) {
    return this.dispatch({ type: 'undo' }, source, options);
  }

  redo(source: ChangeSource = 'ref', options: SpreadsheetDispatchOptions = {}) {
    return this.dispatch({ type: 'redo' }, source, options);
  }

  checkpoint(): SpreadsheetControllerCheckpoint {
    const checkpoint = {
      legacy: this.legacy.checkpoint(),
      documentHistory: this.documentHistory.checkpoint(),
      ['document']: this.checkpointDocument ?? this.currentDocument,
    } as SpreadsheetControllerCheckpoint;
    Object.defineProperty(checkpoint, spreadsheetCheckpointOwner, {
      configurable: false,
      enumerable: false,
      value: this.checkpointOwner,
      writable: false,
    });
    Object.freeze(checkpoint);
    this.checkpoints.add(checkpoint);
    return checkpoint;
  }

  restore(checkpoint: SpreadsheetControllerCheckpoint): void {
    this.assertNoActiveCommitMutation();
    if (
      typeof checkpoint !== 'object' ||
      checkpoint === null ||
      !this.checkpoints.has(checkpoint) ||
      checkpoint[spreadsheetCheckpointOwner] !== this.checkpointOwner
    ) {
      throw new TegoSheetException({
        code: 'INVALID_COMMAND',
        message: 'Checkpoint does not belong to this spreadsheet document controller',
        recoverable: true,
      });
    }
    const rollbackLegacy = this.legacy.checkpoint();
    const rollbackHistory = this.documentHistory.checkpoint();
    const rollbackDocument = this.currentDocument;
    const rollbackFormulaProgram = this.formulaProgram;
    const rollbackFormulaValues = this.formulaValues;
    const { ['document']: checkpointDocument } = checkpoint;
    try {
      const formulaProgram = this.formulaEngine.compile(checkpointDocument);
      const formulaValues = this.formulaEngine.recalculate(
        formulaProgram,
        [],
        calculationEnvironment(
          checkpointDocument,
          checkpoint.legacy.revision,
          this.calculationOptions,
        ),
      ).values;
      this.legacy.restore(checkpoint.legacy);
      this.documentHistory.restore(checkpoint.documentHistory);
      this.currentDocument = checkpointDocument;
      this.formulaProgram = formulaProgram;
      this.formulaValues = formulaValues;
    } catch (error) {
      this.legacy.restore(rollbackLegacy);
      this.documentHistory.restore(rollbackHistory);
      this.currentDocument = rollbackDocument;
      this.formulaProgram = rollbackFormulaProgram;
      this.formulaValues = rollbackFormulaValues;
      throw error;
    }
  }

  replace(input: SpreadsheetDocument): void {
    this.assertNoActiveCommitMutation();
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
    const formulaProgram = this.formulaEngine.compile(parsedDocument);
    const formulaValues = this.formulaEngine.recalculate(
      formulaProgram,
      [],
      calculationEnvironment(parsedDocument, this.getSnapshot().revision, this.calculationOptions),
    ).values;
    const projection = projectDocumentToLegacy(parsedDocument, formulaValues);
    this.legacy.replace(
      projection,
      parsedDocument.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    );
    this.currentDocument = parsedDocument;
    this.formulaProgram = formulaProgram;
    this.formulaValues = formulaValues;
    this.documentHistory.clear();
    this.checkpointOwner = Object.freeze({});
    this.checkpoints = new WeakSet<SpreadsheetControllerCheckpoint>();
  }

  setReadOnly(readOnly: boolean): void {
    this.assertNoActiveCommitMutation();
    this.legacy.setReadOnly(readOnly);
  }

  dispose(): void {
    this.assertNoActiveCommitMutation();
    this.subscriptions.dispose();
    this.legacy.dispose();
  }

  private checkTransaction(
    input: SerializableTransactionEnvelope,
    options: TransactionOptions,
  ): SerializableTransactionEnvelope | TransactionRejection {
    const transaction = snapshotTransaction(input);
    if ('status' in transaction) return transaction;
    if (this.commitMutationActive) {
      return {
        status: 'rejected',
        code: 'COMMAND_NOT_ALLOWED',
        message: 'Commit callback reentrancy is not allowed',
      };
    }
    if (this.permissionGateActive) {
      return {
        status: 'rejected',
        code: 'COMMAND_NOT_ALLOWED',
        message: 'Permission gate reentrancy is not allowed',
      };
    }
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
    if (options.permissionGate !== undefined) {
      if (this.permissionGateActive) {
        return {
          status: 'rejected',
          code: 'COMMAND_NOT_ALLOWED',
          message: 'Permission gate reentrancy is not allowed',
        };
      }
      this.permissionGateActive = true;
      try {
        if (
          options.permissionGate({
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
      } finally {
        this.permissionGateActive = false;
      }
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

  private createDocumentHistoryEntry(
    before: SpreadsheetDocument,
    after: SpreadsheetDocument,
  ): {
    readonly before: DocumentPatch;
    readonly after: DocumentPatch;
    readonly metadata: null;
  } {
    return Object.freeze({
      before: createDocumentPatch(after, before),
      after: createDocumentPatch(before, after),
      metadata: null,
    });
  }

  private applyPatchToDocument(
    document: SpreadsheetDocument,
    patch: DocumentPatch,
  ): SpreadsheetDocument {
    const parsed = parseSpreadsheetDocument(applyDocumentPatch(document, patch));
    if (!parsed.ok) {
      throw new TypeError(
        `Document history patch produced an invalid spreadsheet document: ${JSON.stringify(parsed.diagnostics)}`,
      );
    }
    return parsed.document;
  }

  private assertNoActiveCommitMutation(): void {
    if (!this.commitMutationActive) return;
    throw new TegoSheetException({
      code: 'INVALID_COMMAND',
      message: 'Commit callback cannot reenter the document mutation boundary',
      recoverable: true,
    });
  }
}
