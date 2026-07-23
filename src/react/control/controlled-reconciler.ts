import { type Selection, type SheetId, type TegoSheetError } from '../../core';
import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type {
  SpreadsheetControllerCheckpoint,
  SpreadsheetControllerCommit,
  SpreadsheetDocumentController,
} from '../../core/controller/spreadsheet-document-controller';
import type { DocumentSheetId, SpreadsheetDocument } from '../../document';
import { classifyValueUpdate } from './classify-value-update';
import { createPendingCheckpoint, type PendingCheckpoint } from './pending-checkpoint';

interface AcknowledgedBase {
  readonly key: string;
  readonly checkpoint: SpreadsheetControllerCheckpoint;
  readonly runtimeSheetIds: readonly SheetId[];
}

export interface ReconciliationResult {
  readonly refresh: boolean;
  readonly error?: TegoSheetError;
}

export interface ControlledReconciler {
  readonly getNotificationVersion: () => number;
  readonly record: (commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>) => void;
  readonly reconcile: (value: SpreadsheetDocument) => ReconciliationResult;
}

function replayError(cause?: unknown): TegoSheetError {
  return {
    code: 'INVALID_COMMAND',
    message: 'A pending controlled command could not be replayed',
    recoverable: true,
    ...(cause === undefined ? {} : { cause }),
  };
}

function remapSheet(sheet: SheetId, mapping: ReadonlyMap<SheetId, SheetId>): SheetId {
  return mapping.get(sheet) ?? sheet;
}

function remapDocumentSheet(
  sheet: DocumentSheetId,
  mapping: ReadonlyMap<SheetId, SheetId>,
): DocumentSheetId {
  return (mapping.get(sheet as unknown as SheetId) ?? sheet) as unknown as DocumentSheetId;
}

function remapSelection(selection: Selection, mapping: ReadonlyMap<SheetId, SheetId>): Selection {
  return { ...selection, sheet: remapSheet(selection.sheet, mapping) };
}

function assertNever(command: never): never {
  throw new Error(`Unhandled controlled replay command: ${String(command)}`);
}

export function remapWorkbookCommand(
  command: WorkbookCommand,
  mapping: ReadonlyMap<SheetId, SheetId>,
): WorkbookCommand {
  switch (command.type) {
    case 'set-cell-text':
      return {
        ...command,
        address: {
          ...command.address,
          sheet: remapSheet(command.address.sheet, mapping),
        },
      };
    case 'set-style':
    case 'set-border':
    case 'clear-format':
    case 'clear-contents':
    case 'set-cell-metadata':
    case 'merge':
    case 'unmerge':
    case 'set-filter':
    case 'set-validation':
    case 'remove-validation':
      return { ...command, selection: remapSelection(command.selection, mapping) };
    case 'paint-format':
      return {
        ...command,
        source: remapSelection(command.source, mapping),
        target: remapSelection(command.target, mapping),
      };
    case 'paste-internal':
    case 'autofill':
      return {
        ...command,
        source: remapSelection(command.source, mapping),
        target: remapSelection(command.target, mapping),
      };
    case 'paste-external':
      return { ...command, target: remapSelection(command.target, mapping) };
    case 'insert-row':
    case 'delete-row':
    case 'insert-column':
    case 'delete-column':
    case 'set-row-height':
    case 'set-row-hidden':
    case 'set-column-width':
    case 'set-column-hidden':
    case 'set-freeze':
    case 'delete-sheet':
    case 'rename-sheet':
    case 'clear-filter':
    case 'sort':
    case 'remove-filter-view':
    case 'remove-conditional-format':
    case 'remove-sheet-object':
      return { ...command, sheet: remapSheet(command.sheet, mapping) };
    case 'set-filter-view':
      return {
        ...command,
        sheet: remapSheet(command.sheet, mapping),
        view: {
          ...command.view,
          range: {
            ...command.view.range,
            sheetId: remapDocumentSheet(command.view.range.sheetId, mapping),
          },
        },
      };
    case 'set-conditional-format':
      return {
        ...command,
        sheet: remapSheet(command.sheet, mapping),
        format: {
          ...command.format,
          range: {
            ...command.format.range,
            sheetId: remapDocumentSheet(command.format.range.sheetId, mapping),
          },
        },
      };
    case 'set-sheet-object':
      return {
        ...command,
        sheet: remapSheet(command.sheet, mapping),
        object: {
          ...command.object,
          anchor:
            command.object.anchor.type === 'absolute'
              ? command.object.anchor
              : command.object.anchor.type === 'one-cell'
                ? {
                    ...command.object.anchor,
                    cell: {
                      ...command.object.anchor.cell,
                      sheetId: remapDocumentSheet(command.object.anchor.cell.sheetId, mapping),
                    },
                  }
                : {
                    ...command.object.anchor,
                    from: {
                      ...command.object.anchor.from,
                      sheetId: remapDocumentSheet(command.object.anchor.from.sheetId, mapping),
                    },
                    to: {
                      ...command.object.anchor.to,
                      sheetId: remapDocumentSheet(command.object.anchor.to.sheetId, mapping),
                    },
                  },
        },
      };
    case 'add-sheet':
    case 'undo':
    case 'redo':
      return command;
    default:
      return assertNever(command);
  }
}

function mapSheetIds(
  mapping: Map<SheetId, SheetId>,
  original: readonly SheetId[],
  replayed: readonly SheetId[],
): void {
  const count = Math.min(original.length, replayed.length);
  for (let index = 0; index < count; index += 1) {
    mapping.set(original[index]!, replayed[index]!);
  }
}

function sameSheetIds(actual: readonly SheetId[], expected: readonly SheetId[]): boolean {
  return (
    actual.length === expected.length && actual.every((sheet, index) => sheet === expected[index])
  );
}

function reportedReference(
  weak: WeakSet<object>,
  primitive: Set<unknown>,
  value: unknown,
): boolean {
  if ((typeof value === 'object' && value !== null) || typeof value === 'function') {
    if (weak.has(value)) return true;
    weak.add(value);
    return false;
  }
  if (primitive.has(value)) return true;
  primitive.add(value);
  return false;
}

export function createControlledReconciler(
  controller: SpreadsheetDocumentController,
): ControlledReconciler {
  const initialDocument = controller.getDocument();
  const neverObserved = Symbol('controlled-value-not-observed');
  let observedValue: unknown = neverObserved;
  let pending: PendingCheckpoint[] = [];
  let base: AcknowledgedBase = {
    key: JSON.stringify(initialDocument),
    checkpoint: controller.checkpoint(),
    runtimeSheetIds: controller.getSheetIds(),
  };
  let notificationVersion = 0;
  const invalidObjects = new WeakSet<object>();
  const invalidPrimitives = new Set<unknown>();

  const replay = (
    tail: readonly PendingCheckpoint[],
  ): {
    readonly pending: PendingCheckpoint[];
    readonly error?: TegoSheetError;
  } => {
    const replayed: PendingCheckpoint[] = [];
    const mapping = new Map<SheetId, SheetId>();
    mapSheetIds(mapping, base.runtimeSheetIds, controller.getSheetIds());
    const readOnly = controller.getSnapshot().readOnly;
    if (readOnly) controller.setReadOnly(false);
    try {
      for (const original of tail) {
        const before = controller.checkpoint();
        let outcome:
          | { readonly status: 'noop' }
          | {
              readonly status: 'committed';
              readonly commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>;
            };
        try {
          outcome = controller.dispatch(
            remapWorkbookCommand(original.command, mapping),
            original.source,
            {
              notify: false,
              replayAddSheetId: original.addedSheetId,
            },
          );
        } catch (cause) {
          controller.restore(before);
          return { pending: replayed, error: replayError(cause) };
        }
        const replayedSheetIds = controller.getSheetIds();
        if (
          outcome.status !== 'committed' ||
          JSON.stringify(outcome.commit.document) !== original.projectedKey ||
          !sameSheetIds(replayedSheetIds, original.runtimeSheetIds)
        ) {
          controller.restore(before);
          return { pending: replayed, error: replayError() };
        }
        mapSheetIds(mapping, original.runtimeSheetIds, replayedSheetIds);
        replayed.push(createPendingCheckpoint(controller, outcome.commit, original));
      }
      return { pending: replayed };
    } finally {
      if (readOnly) controller.setReadOnly(true);
    }
  };

  return {
    getNotificationVersion: () => notificationVersion,
    record(commit) {
      pending.push(createPendingCheckpoint(controller, commit));
    },
    reconcile(value) {
      const update = classifyValueUpdate(
        {
          observedValue,
          acknowledgedKey: base.key,
          pending,
        },
        value,
      );
      if (update.kind === 'same-reference') {
        return { refresh: false };
      }
      observedValue = value;
      if (update.kind === 'invalid') {
        const duplicate = reportedReference(invalidObjects, invalidPrimitives, value);
        return {
          refresh: false,
          ...(duplicate ? {} : { error: update.error }),
        };
      }
      if (update.kind === 'replace') {
        controller.replace(update.document);
        base = {
          key: JSON.stringify(update.document),
          checkpoint: controller.checkpoint(),
          runtimeSheetIds: controller.getSheetIds(),
        };
        pending = [];
        notificationVersion += 1;
        return { refresh: true };
      }
      if (update.kind === 'rollback') {
        if (pending.length === 0) return { refresh: false };
        controller.restore(base.checkpoint);
        pending = [];
        notificationVersion += 1;
        return { refresh: true };
      }

      const acknowledged = pending[update.through]!;
      const tail = pending.slice(update.through + 1);
      controller.restore(acknowledged.checkpoint);
      base = {
        key: acknowledged.projectedKey,
        checkpoint: acknowledged.checkpoint,
        runtimeSheetIds: acknowledged.runtimeSheetIds,
      };
      const replayed = replay(tail);
      pending = replayed.pending;
      if (replayed.error !== undefined) notificationVersion += 1;
      return {
        refresh: true,
        ...(replayed.error === undefined ? {} : { error: replayed.error }),
      };
    },
  };
}
