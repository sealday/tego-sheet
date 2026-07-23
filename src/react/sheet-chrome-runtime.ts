import type { Selection, SheetData, SheetId } from '../core';
import type { SpreadsheetControllerSnapshot } from '../core/controller/spreadsheet-document-controller';
import { filterItems, parseA1Range } from '../core';

export function activeSheetData(
  snapshot: SpreadsheetControllerSnapshot,
  activeSheet: SheetId | null,
): SheetData | null {
  const index = snapshot.sheets.findIndex((sheet) => sheet.id === activeSheet);
  return index < 0 ? null : (snapshot.projection[index] ?? null);
}

export function filterValuesForSelection(
  sheet: SheetData,
  selection: Selection,
): readonly string[] {
  let range = selection.range;
  if (sheet.autofilter?.ref !== undefined) {
    try {
      const filterRange = parseA1Range(sheet.autofilter.ref);
      if (
        selection.active.column >= filterRange.start.column &&
        selection.active.column <= filterRange.end.column
      )
        range = filterRange;
    } catch {
      // Invalid imported filter references remain inert until replaced by a valid command.
    }
  }
  return Object.keys(filterItems(sheet, selection.active.column, range));
}

export function filterCommandSelection(sheet: SheetData, selection: Selection): Selection {
  if (sheet.autofilter?.ref === undefined) return selection;
  try {
    const range = parseA1Range(sheet.autofilter.ref);
    if (selection.active.column < range.start.column || selection.active.column > range.end.column)
      return selection;
    return {
      ...selection,
      active: { row: range.start.row, column: selection.active.column },
      range,
    };
  } catch {
    return selection;
  }
}
