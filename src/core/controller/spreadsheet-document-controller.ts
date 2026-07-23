import { parseSpreadsheetDocument } from '../../document/parse-document';
import { projectDocumentToLegacy, projectLegacyToDocument } from './runtime-projection';
import type { SpreadsheetDocument } from '../../document/model/document';
import type { CommandResult, WorkbookCommand } from '../commands/workbook-command';
import type { ChangeSource } from '../types/changes';
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
}
