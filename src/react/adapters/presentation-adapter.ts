import { validateWorkbook } from '../../core';
import type { SpreadsheetControllerSnapshot } from '../../core/controller/spreadsheet-document-controller';
import type { DocumentCellAddress } from '../../document';
import type { PresentationValidation } from '../../presentation';

/** Builds one immutable validation lookup for a controller snapshot. */
export function createPresentationValidationResolver(
  snapshot: SpreadsheetControllerSnapshot,
): (address: DocumentCellAddress) => PresentationValidation {
  const result = validateWorkbook({
    getValue: () => snapshot.projection,
    getSheetIds: () => snapshot.sheets.map(({ id }) => id),
  });
  const issues = new Map(
    result.issues.map((issue) => [
      `${issue.sheet}:${issue.address.row}:${issue.address.column}`,
      issue.message,
    ]),
  );
  return (address) => {
    const message = issues.get(`${address.sheetId}:${address.row}:${address.column}`);
    return message === undefined ? { status: 'valid' } : { status: 'error', message };
  };
}
