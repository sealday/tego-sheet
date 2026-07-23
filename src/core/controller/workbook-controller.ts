import { applyCommand } from '../commands/apply-command';
import type { CommandCommit, CommandOutcome } from '../commands/command-result';
import { invalidCommand, validateCommand } from '../commands/validate-command';
import type { CommandResult, WorkbookCommand } from '../commands/workbook-command';
import { WorkbookState } from '../model/workbook-state';
import { selectCellText } from '../selectors/cell';
import type { ChangeSource, WorkbookChange, WorkbookChangeKind } from '../types/changes';
import { assertCellAddress } from '../types/coordinates';
import type { CellAddress, SheetId } from '../types/coordinates';
import type { WorkbookData, WorkbookInput } from '../types/workbook';
import type { ValidationResult } from '../types/validation';
import { validateWorkbook } from '../selectors/validation';
import {
  createControllerCheckpoint,
  hasCheckpointOwner,
  type ControllerCheckpoint,
  type HistoryMetadata,
  type WorkbookPatch,
} from './controller-checkpoint';
import { History, type HistoryCheckpoint, type HistoryEntry } from './history';
import { SubscriptionStore } from './subscription-store';
import type { WorkbookInitializationDefaults } from '../serialization/canonicalize-workbook';
import { applyDocumentPatch, createDocumentPatch } from './document-patch';

export interface WorkbookControllerOptions {
  readonly readOnly?: boolean;
  readonly initialRowCount?: number;
  readonly initialColumnCount?: number;
  /** @internal Stable IDs supplied by the schema 2 controller projection. */
  readonly sheetIds?: readonly SheetId[];
}

export interface DispatchOptions {
  /**
   * Run an internal atomic observer after commit construction and before subscriptions publish.
   * If it throws, state, history, revision, and change sequencing are rolled back before the
   * original exception is rethrown, and no subscription is published.
   */
  readonly beforeNotify?: (commit: CommandCommit<unknown, WorkbookCommand>) => void;
  /** Suppress the document subscription used by controlled replay and restore. */
  readonly notify?: boolean;
  /** Skip the potentially large paste result when no consumer needs it. */
  readonly capturePasteValues?: boolean;
  /** Preserve the originally exposed runtime ID while replaying a pending add-sheet command. */
  readonly replayAddSheetId?: SheetId;
}

export interface ControllerSheetSnapshot {
  readonly id: SheetId;
  readonly index: number;
  readonly name: string;
}

export interface ControllerSnapshot {
  readonly revision: number;
  readonly value: WorkbookData;
  readonly sheets: readonly ControllerSheetSnapshot[];
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly readOnly: boolean;
}

export interface ControllerEvent {
  readonly snapshot: ControllerSnapshot;
  readonly commit: CommandCommit<unknown, WorkbookCommand>;
}

export type ControllerSubscriber = (event: ControllerEvent) => void;

interface MutationTransaction {
  readonly state: WorkbookState;
  readonly history: HistoryCheckpoint<WorkbookPatch, HistoryMetadata>;
  readonly revision: number;
  readonly changeSequence: number;
}

export interface ProjectionCommit<Result> {
  readonly value: WorkbookInput;
  readonly sheetIds: readonly SheetId[];
  readonly result: Result;
  readonly kind: WorkbookChangeKind;
  readonly sheet: SheetId;
  readonly range?: WorkbookChange['range'];
}

export interface TransactionChangeSummary {
  readonly kind: WorkbookChangeKind;
  readonly sheet: SheetId;
  readonly range?: WorkbookChange['range'];
  readonly aggregate?: WorkbookChange['aggregate'];
}

let nextControllerId = 1;

function cloneValue<T>(value: T): T {
  if (Array.isArray(value)) return value.map((item) => cloneValue(item)) as T;
  if (value !== null && typeof value === 'object') {
    const output: Record<string, unknown> = {};
    for (const key of Object.keys(value)) {
      Object.defineProperty(output, key, {
        configurable: true,
        enumerable: true,
        value: cloneValue((value as Record<string, unknown>)[key]),
        writable: true,
      });
    }
    return output as T;
  }
  return value;
}

function freezeValue<T>(value: T): T {
  if (value !== null && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const key of Object.keys(value)) {
      freezeValue((value as Record<string, unknown>)[key]);
    }
    Object.freeze(value);
  }
  return value;
}

function isolated<T>(value: T): T {
  return freezeValue(cloneValue(value));
}

export class WorkbookController {
  private state: WorkbookState;
  private readonly history = new History<WorkbookPatch, HistoryMetadata>();
  private readonly subscriptions = new SubscriptionStore<ControllerEvent>();
  private checkpointOwner: object = Object.freeze({});
  private checkpoints = new WeakSet<ControllerCheckpoint>();
  private readonly controllerId: number;
  private revision = 0;
  private changeSequence = 0;
  private readOnly: boolean;
  private disposed = false;
  private readonly initializationDefaults: Readonly<WorkbookInitializationDefaults>;

  constructor(input: WorkbookInput, options: WorkbookControllerOptions = {}) {
    this.initializationDefaults = Object.freeze({
      rowCount: options.initialRowCount,
      columnCount: options.initialColumnCount,
    });
    this.state = WorkbookState.from(input, this.initializationDefaults, options.sheetIds);
    this.readOnly = options.readOnly ?? false;
    this.controllerId = nextControllerId;
    nextControllerId += 1;
  }

  get historySize(): { readonly undo: number; readonly redo: number } {
    return this.history.size;
  }

  get canUndo(): boolean {
    return this.history.canUndo;
  }

  get canRedo(): boolean {
    return this.history.canRedo;
  }

  getValue(): WorkbookData {
    return this.state.serialize();
  }

  getSheetIds(): readonly SheetId[] {
    return Object.freeze(this.state.sheets.map((sheet) => sheet.id));
  }

  getInitializationDefaults(): Readonly<WorkbookInitializationDefaults> {
    return this.initializationDefaults;
  }

  getCellText(address: CellAddress): string {
    try {
      assertCellAddress(address);
    } catch (cause) {
      throw invalidCommand('Cell query requires a valid address', cause);
    }
    const sheet = this.state.get(address.sheet);
    if (sheet === null) throw invalidCommand(`Unknown sheet ID: ${address.sheet}`);
    return selectCellText(sheet.data, address.row, address.column);
  }

  getSnapshot(): ControllerSnapshot {
    const value = this.getValue();
    return isolated({
      revision: this.revision,
      value,
      sheets: this.state.sheets.map((sheet, index) => ({
        id: sheet.id,
        index,
        name: sheet.data.name ?? '',
      })),
      canUndo: this.history.canUndo,
      canRedo: this.history.canRedo,
      readOnly: this.readOnly,
    });
  }

  validate(): ValidationResult {
    this.ensureActive();
    return validateWorkbook(this);
  }

  subscribe(subscriber: ControllerSubscriber): () => void {
    this.ensureActive();
    return this.subscriptions.subscribe(subscriber);
  }

  /** @internal Validates a command before an outer controller prepares transaction provenance. */
  assertCommand(command: WorkbookCommand): void {
    this.ensureMutable();
    validateCommand(this.state, this.isolateCommand(command));
  }

  dispatch<Command extends WorkbookCommand>(
    command: Command,
    source: ChangeSource,
    options: DispatchOptions = {},
  ): CommandOutcome<CommandResult<Command>, Command> {
    this.ensureMutable();
    const commandSnapshot = this.isolateCommand(command);
    if (
      options.replayAddSheetId !== undefined &&
      (commandSnapshot.type !== 'add-sheet' || options.notify !== false)
    ) {
      throw invalidCommand('A replay sheet ID requires a silent add-sheet command');
    }
    validateCommand(this.state, commandSnapshot);
    if (commandSnapshot.type === 'undo') {
      return this.applyHistory('undo', commandSnapshot, source, options) as CommandOutcome<
        CommandResult<Command>,
        Command
      >;
    }
    if (commandSnapshot.type === 'redo') {
      return this.applyHistory('redo', commandSnapshot, source, options) as CommandOutcome<
        CommandResult<Command>,
        Command
      >;
    }

    const applied = applyCommand(this.state, commandSnapshot, {
      capturePasteValues: options.capturePasteValues !== false,
      replayAddSheetId: options.replayAddSheetId,
    });
    if (applied === null) return { status: 'noop' };

    const transaction = options.beforeNotify === undefined ? undefined : this.captureTransaction();
    const before = this.state;
    const change = this.createChange(applied.kind, source, applied.sheet, applied.range);
    this.state = applied.state;
    this.revision += 1;
    if (applied.undoable) {
      const metadata = Object.freeze<HistoryMetadata>({ command: commandSnapshot, change });
      this.history.record(this.createHistoryEntry(before, applied.state, metadata));
    }

    const commit = this.createCommit(
      commandSnapshot,
      change,
      applied.result as CommandResult<Command>,
    );
    this.runBeforeNotify(commit as CommandCommit<unknown, WorkbookCommand>, options, transaction);
    this.publish(commit, options);
    return { status: 'committed', commit };
  }

  /** @internal Commits a schema-authored projection when the legacy operation was a no-op. */
  commitProjection<Command extends WorkbookCommand>(
    command: Command,
    source: ChangeSource,
    projection: ProjectionCommit<CommandResult<Command>>,
    options: DispatchOptions = {},
  ): CommandOutcome<CommandResult<Command>, Command> {
    this.ensureMutable();
    const commandSnapshot = this.isolateCommand(command);
    validateCommand(this.state, commandSnapshot);
    const transaction = options.beforeNotify === undefined ? undefined : this.captureTransaction();
    const before = this.state;
    const after = before.replace(projection.value, projection.sheetIds);
    const change = this.createChange(projection.kind, source, projection.sheet, projection.range);
    this.state = after;
    this.revision += 1;
    const metadata = Object.freeze<HistoryMetadata>({ command: commandSnapshot, change });
    this.history.record(this.createHistoryEntry(before, after, metadata));
    const commit = this.createCommit(commandSnapshot, change, projection.result);
    this.runBeforeNotify(commit as CommandCommit<unknown, WorkbookCommand>, options, transaction);
    this.publish(commit, options);
    return { status: 'committed', commit };
  }

  /**
   * @internal Reconciles schema-authored fields into the current operational state and its
   * matching history entry. This is called from an active `beforeNotify` transaction.
   */
  reconcileProjection(
    input: WorkbookInput,
    sheetIds: readonly SheetId[],
    direction: 'commit' | 'undo' | 'redo',
  ): void {
    this.ensureActive();
    const state = this.state.replace(input, sheetIds);
    const operationalState = this.state;
    const checkpoint = this.history.checkpoint();
    if (direction === 'undo') {
      const redo = [...checkpoint.redo];
      const entry = redo.at(-1);
      if (entry !== undefined) {
        const after = this.applyPatchToState(operationalState, entry.after);
        redo[redo.length - 1] = this.createHistoryEntry(state, after, entry.metadata);
      }
      this.history.restore({ undo: checkpoint.undo, redo });
    } else {
      const undo = [...checkpoint.undo];
      const entry = undo.at(-1);
      if (entry !== undefined) {
        const before = this.applyPatchToState(operationalState, entry.before);
        undo[undo.length - 1] = this.createHistoryEntry(before, state, entry.metadata);
      }
      this.history.restore({ undo, redo: checkpoint.redo });
    }
    this.state = state;
  }

  undo(
    source: ChangeSource = 'ref',
    options: DispatchOptions = {},
  ): CommandOutcome<void, { readonly type: 'undo' }> {
    return this.dispatch({ type: 'undo' }, source, options);
  }

  redo(
    source: ChangeSource = 'ref',
    options: DispatchOptions = {},
  ): CommandOutcome<void, { readonly type: 'redo' }> {
    return this.dispatch({ type: 'redo' }, source, options);
  }

  checkpoint(): ControllerCheckpoint {
    this.ensureActive();
    const checkpoint = createControllerCheckpoint(
      this.state,
      this.history.checkpoint(),
      this.revision,
      this.changeSequence,
      this.checkpointOwner,
    );
    this.checkpoints.add(checkpoint);
    return checkpoint;
  }

  restore(checkpoint: ControllerCheckpoint): void {
    this.ensureActive();
    if (
      typeof checkpoint !== 'object' ||
      checkpoint === null ||
      !this.checkpoints.has(checkpoint) ||
      !hasCheckpointOwner(checkpoint, this.checkpointOwner)
    ) {
      throw invalidCommand('Checkpoint does not belong to this workbook controller');
    }
    this.state = checkpoint.state;
    this.history.restore(checkpoint.history);
    this.revision = checkpoint.revision;
    this.changeSequence = checkpoint.changeSequence;
  }

  /**
   * @internal Collapses silently prepared commands into one revision and one history entry.
   * The checkpoint must belong to this controller and describe the state before preparation.
   */
  finalizeTransaction(
    checkpoint: ControllerCheckpoint,
    command: WorkbookCommand,
    source: ChangeSource,
    summary: TransactionChangeSummary,
    forceChanged = false,
  ): CommandOutcome<void, WorkbookCommand> {
    this.ensureMutable();
    if (!this.isOwnedCheckpoint(checkpoint)) {
      throw invalidCommand('Checkpoint does not belong to this workbook controller');
    }
    const after = this.state;
    const changed =
      JSON.stringify(after.serialize()) !== JSON.stringify(checkpoint.state.serialize());
    this.history.restore(checkpoint.history);
    this.revision = checkpoint.revision;
    this.changeSequence = checkpoint.changeSequence;
    if (!changed && !forceChanged) {
      this.state = checkpoint.state;
      return { status: 'noop' };
    }
    this.state = after;
    const commandSnapshot = this.isolateCommand(command);
    const change = this.createChange(
      summary.kind,
      source,
      summary.sheet,
      summary.range,
      summary.aggregate,
    );
    this.revision += 1;
    this.history.record(
      this.createHistoryEntry(
        checkpoint.state,
        after,
        Object.freeze({ command: commandSnapshot, change }),
      ),
    );
    return {
      status: 'committed',
      commit: this.createCommit(commandSnapshot, change, undefined),
    };
  }

  replace(input: WorkbookInput, sheetIds?: readonly SheetId[]): void {
    this.ensureActive();
    const replacement = this.state.replace(input, sheetIds);
    this.state = replacement;
    this.history.clear();
    this.revision += 1;
    this.checkpointOwner = Object.freeze({});
    this.checkpoints = new WeakSet<ControllerCheckpoint>();
  }

  setReadOnly(readOnly: boolean): void {
    this.ensureActive();
    this.readOnly = readOnly;
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.subscriptions.dispose();
    this.history.clear();
  }

  private applyHistory<Command extends { readonly type: 'undo' | 'redo' }>(
    direction: 'undo' | 'redo',
    command: Command,
    source: ChangeSource,
    options: DispatchOptions,
  ): CommandOutcome<void, Command> {
    const transaction = options.beforeNotify === undefined ? undefined : this.captureTransaction();
    const entry = direction === 'undo' ? this.history.undo() : this.history.redo();
    if (entry === null) return { status: 'noop' };
    this.state = this.applyPatchToState(
      this.state,
      direction === 'undo' ? entry.before : entry.after,
    );
    this.revision += 1;
    const change = this.createChange(
      'history',
      source,
      entry.metadata.change.sheet,
      entry.metadata.change.range,
      entry.metadata.change.aggregate,
    );
    const commit = this.createCommit(command, change, undefined);
    this.runBeforeNotify(commit as CommandCommit<unknown, WorkbookCommand>, options, transaction);
    this.publish(commit, options);
    return { status: 'committed', commit };
  }

  private createChange(
    kind: WorkbookChange['kind'],
    source: ChangeSource,
    sheet: SheetId,
    range?: WorkbookChange['range'],
    aggregate?: WorkbookChange['aggregate'],
  ): WorkbookChange {
    this.changeSequence += 1;
    return isolated({
      id: `change-${this.controllerId}-${this.changeSequence}`,
      kind,
      source,
      sheet,
      ...(range === undefined ? {} : { range }),
      ...(aggregate === undefined ? {} : { aggregate }),
    });
  }

  private createPatch(before: WorkbookState, after: WorkbookState): WorkbookPatch {
    return Object.freeze({
      ...createDocumentPatch(before.serialize(), after.serialize()),
      sheetIds: Object.freeze(after.sheets.map((sheet) => sheet.id)),
    });
  }

  private createHistoryEntry(
    before: WorkbookState,
    after: WorkbookState,
    metadata: HistoryMetadata,
  ): HistoryEntry<WorkbookPatch, HistoryMetadata> {
    return Object.freeze({
      before: this.createPatch(after, before),
      after: this.createPatch(before, after),
      metadata,
    });
  }

  private applyPatchToState(state: WorkbookState, patch: WorkbookPatch): WorkbookState {
    return state.replace(applyDocumentPatch(state.serialize(), patch), patch.sheetIds);
  }

  private captureTransaction(): MutationTransaction {
    return {
      state: this.state,
      history: this.history.checkpoint(),
      revision: this.revision,
      changeSequence: this.changeSequence,
    };
  }

  private runBeforeNotify(
    commit: CommandCommit<unknown, WorkbookCommand>,
    options: DispatchOptions,
    transaction: MutationTransaction | undefined,
  ): void {
    const beforeNotify = options.beforeNotify;
    if (beforeNotify === undefined) return;
    if (transaction === undefined) throw new Error('Missing beforeNotify transaction');
    try {
      beforeNotify(commit);
    } catch (error) {
      this.state = transaction.state;
      this.history.restore(transaction.history);
      this.revision = transaction.revision;
      this.changeSequence = transaction.changeSequence;
      throw error;
    }
  }

  private createCommit<Result, Command extends WorkbookCommand>(
    command: Command,
    change: WorkbookChange,
    result: Result,
  ): CommandCommit<Result, Command> {
    return Object.freeze({
      command,
      change,
      result: isolated(result),
      value: isolated(this.getValue()),
    });
  }

  private isolateCommand<Command extends WorkbookCommand>(command: Command): Command {
    try {
      const snapshot = isolated(command);
      if (snapshot === null || typeof snapshot !== 'object' || Array.isArray(snapshot)) {
        throw new TypeError('Command must be an object');
      }
      return snapshot;
    } catch (cause) {
      throw invalidCommand('Command could not be isolated', cause);
    }
  }

  private publish<Result, Command extends WorkbookCommand>(
    commit: CommandCommit<Result, Command>,
    options: DispatchOptions,
  ): void {
    if (options.notify === false) return;
    this.subscriptions.publish(
      Object.freeze({
        snapshot: this.getSnapshot(),
        commit: commit as CommandCommit<unknown, WorkbookCommand>,
      }),
    );
  }

  private ensureMutable(): void {
    this.ensureActive();
    if (this.readOnly) throw invalidCommand('Workbook is read-only');
  }

  private ensureActive(): void {
    if (this.disposed) throw invalidCommand('Workbook controller is disposed');
  }

  private isOwnedCheckpoint(checkpoint: ControllerCheckpoint): boolean {
    return (
      typeof checkpoint === 'object' &&
      checkpoint !== null &&
      this.checkpoints.has(checkpoint) &&
      hasCheckpointOwner(checkpoint, this.checkpointOwner)
    );
  }
}

export type WorkbookHistoryEntry = HistoryEntry<WorkbookPatch, HistoryMetadata>;
