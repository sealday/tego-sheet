import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type {
  SpreadsheetControllerCheckpoint,
  SpreadsheetControllerCommit,
  SpreadsheetDocumentController,
} from '../../core/controller/spreadsheet-document-controller';
import type { ChangeSource, SheetId } from '../../core';
import type { SpreadsheetDocument } from '../../document';

export interface PendingCheckpoint {
  readonly command: WorkbookCommand;
  readonly source: ChangeSource;
  readonly changeId: string;
  readonly projected: SpreadsheetDocument;
  readonly projectedKey: string;
  readonly checkpoint: SpreadsheetControllerCheckpoint;
  readonly runtimeSheetIds: readonly SheetId[];
  readonly addedSheetId?: SheetId;
}

export function createPendingCheckpoint(
  controller: SpreadsheetDocumentController,
  commit: SpreadsheetControllerCommit<unknown, WorkbookCommand>,
  original?: PendingCheckpoint,
): PendingCheckpoint {
  const addedSheetId =
    original?.addedSheetId ??
    (commit.command.type === 'add-sheet' && typeof commit.result === 'string'
      ? (commit.result as SheetId)
      : undefined);
  return Object.freeze({
    command: original?.command ?? commit.command,
    source: original?.source ?? commit.change.source,
    changeId: original?.changeId ?? commit.change.id,
    projected: original?.projected ?? commit.document,
    projectedKey: original?.projectedKey ?? JSON.stringify(commit.document),
    checkpoint: controller.checkpoint(),
    runtimeSheetIds: original?.runtimeSheetIds ?? controller.getSheetIds(),
    ...(addedSheetId === undefined ? {} : { addedSheetId }),
  });
}
