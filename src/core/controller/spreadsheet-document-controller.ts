import { parseSpreadsheetDocument } from '../../document/parse-document';
import {
  projectDocumentToLegacy,
  projectLegacyToDocument,
} from '../../document/runtime-projection';
import type { SpreadsheetDocument } from '../../document/model/document';
import type { CommandResult, WorkbookCommand } from '../commands/workbook-command';
import type { ChangeSource } from '../types/changes';
import { sheetId, type CellAddress, type SheetId } from '../types/coordinates';
import type { ValidationResult } from '../types/validation';
import {
  WorkbookController,
  type ControllerSnapshot,
  type DispatchOptions,
  type WorkbookControllerOptions,
} from './workbook-controller';

export interface SpreadsheetControllerCommit<
  Result = void,
  Command extends WorkbookCommand = WorkbookCommand,
> {
  readonly command: Command;
  readonly change: import('../types/changes').WorkbookChange;
  readonly result: Result;
  readonly document: SpreadsheetDocument;
}

export interface SpreadsheetControllerSnapshot extends Omit<ControllerSnapshot, 'value'> {
  readonly document: SpreadsheetDocument;
  /** Read-only projection consumed only by the current engine boundary. */
  readonly projection: ControllerSnapshot['value'];
}

export interface SpreadsheetControllerEvent {
  readonly snapshot: SpreadsheetControllerSnapshot;
  readonly commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>;
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
  private document: SpreadsheetDocument;
  private readonly legacy: WorkbookController;
  private readonly subscribers = new Set<(event: SpreadsheetControllerEvent) => void>();

  constructor(input: SpreadsheetDocument, options: WorkbookControllerOptions = {}) {
    const parsed = parseSpreadsheetDocument(input);
    if (!parsed.ok) throw new TypeError('A valid spreadsheet document is required');
    this.document = parsed.document;
    this.legacy = new WorkbookController(projectDocumentToLegacy(this.document), {
      ...options,
      sheetIds: this.document.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    });
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
    return cloneFrozen(this.document);
  }

  getValue() {
    return this.legacy.getValue();
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
    return cloneFrozen({
      ...snapshot,
      document: this.document,
      projection: snapshot.value,
      value: undefined,
    }) as SpreadsheetControllerSnapshot;
  }

  validate(): ValidationResult {
    return this.legacy.validate();
  }

  subscribe(subscriber: (event: SpreadsheetControllerEvent) => void): () => void {
    this.subscribers.add(subscriber);
    return () => this.subscribers.delete(subscriber);
  }

  dispatch<Command extends WorkbookCommand>(
    command: Command,
    source: ChangeSource,
    options: DispatchOptions = {},
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
      document: this.document,
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
      for (const subscriber of this.subscribers) subscriber(event);
    }
    return { status: 'committed', commit };
  }

  undo(source: ChangeSource = 'ref', options: DispatchOptions = {}) {
    return this.dispatch({ type: 'undo' }, source, options);
  }

  redo(source: ChangeSource = 'ref', options: DispatchOptions = {}) {
    return this.dispatch({ type: 'redo' }, source, options);
  }

  checkpoint() {
    return {
      legacy: this.legacy.checkpoint(),
      document: this.document,
    };
  }

  restore(checkpoint: ReturnType<SpreadsheetDocumentController['checkpoint']>): void {
    this.legacy.restore(checkpoint.legacy);
    this.document = checkpoint.document;
  }

  replace(input: SpreadsheetDocument): void {
    const parsed = parseSpreadsheetDocument(input);
    if (!parsed.ok) throw new TypeError('A valid spreadsheet document is required');
    this.document = parsed.document;
    this.legacy.replace(
      projectDocumentToLegacy(this.document),
      this.document.workbook.sheets.map((sheet) => sheetId(sheet.id)),
    );
  }

  setReadOnly(readOnly: boolean): void {
    this.legacy.setReadOnly(readOnly);
  }

  dispose(): void {
    this.subscribers.clear();
    this.legacy.dispose();
  }

  private refreshDocument(): void {
    this.document = projectLegacyToDocument(
      this.legacy.getValue(),
      this.document,
      this.legacy.getSheetIds(),
    );
  }
}
