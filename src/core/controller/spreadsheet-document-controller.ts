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
  readonly ['document']: SpreadsheetDocument;
}

function cloneFrozen<T>(value: T): T {
  if (value === null || typeof value !== 'object') return value;
  const output = (Array.isArray(value) ? [] : {}) as Record<string, unknown>;
  for (const [key, item] of Object.entries(value)) output[key] = cloneFrozen(item);
  return Object.freeze(output) as T;
}

/**
 * Owns the single schema 2 runtime truth while adapting existing operations at one private boundary.
 */
export class SpreadsheetDocumentController {
  private currentDocument: SpreadsheetDocument;
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
    this.refreshDocument();
  }

  get historySize() {
    return this.legacy.historySize;
  }

  get canUndo(): boolean {
    return this.legacy.canUndo;
  }

  get canRedo(): boolean {
    return this.legacy.canRedo;
  }

  getDocument(): SpreadsheetDocument {
    return cloneFrozen(this.currentDocument);
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
    return cloneFrozen({
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
    const checkpoint = options.beforeNotify === undefined ? undefined : this.checkpoint();
    const outcome = this.legacy.dispatch(command, source, {
      ...options,
      beforeNotify: undefined,
      notify: false,
    });
    if (outcome.status === 'noop') return outcome;
    this.refreshDocument();
    const commit = cloneFrozen({
      command: outcome.commit.command,
      change: outcome.commit.change,
      result: outcome.commit.result,
      ['document']: this.currentDocument,
    }) as SpreadsheetControllerCommit<CommandResult<Command>, Command>;
    try {
      options.beforeNotify?.(commit as never);
    } catch (error) {
      if (checkpoint !== undefined) this.restore(checkpoint);
      throw error;
    }
    if (options.notify !== false) {
      const event = cloneFrozen({
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
      ['document']: this.currentDocument,
    };
  }

  restore(checkpoint: SpreadsheetControllerCheckpoint): void {
    this.legacy.restore(checkpoint.legacy);
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
    this.currentDocument = parsedDocument;
    this.legacy.replace(
      projectDocumentToLegacy(this.currentDocument),
      this.currentDocument.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    );
  }

  setReadOnly(readOnly: boolean): void {
    this.legacy.setReadOnly(readOnly);
  }

  dispose(): void {
    this.subscriptions.dispose();
    this.legacy.dispose();
  }

  private refreshDocument(): void {
    this.currentDocument = projectLegacyToDocument(
      this.legacy.getValue(),
      this.currentDocument,
      this.legacy.getSheetIds(),
    );
  }
}
