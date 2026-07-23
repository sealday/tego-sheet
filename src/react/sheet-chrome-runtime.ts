import type { CellStyle, Selection, SheetData, SheetId } from '../core';
import type { SpreadsheetControllerSnapshot } from '../core/controller/spreadsheet-document-controller';
import { filterItems, parseA1Range } from '../core';
import { createPrintLayout } from '../engine';
import {
  createFontMetrics,
  createPresentationCache,
  createPresentationResolver,
} from '../presentation';
import { createPrintDisplayList } from '../print';
import { mountPrintDisplayPages, type PrintWorkbookOptions } from '../ui/print-workbook';

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

export function mountActiveSheetPrint(
  snapshot: SpreadsheetControllerSnapshot,
  activeSheet: SheetId | null,
  options: PrintWorkbookOptions,
  defaultStyle?: CellStyle,
): (() => void) | null {
  const sheet = activeSheetData(snapshot, activeSheet);
  const sheetIndex = snapshot.sheets.findIndex(({ id }) => id === activeSheet);
  const documentSheet = sheetIndex < 0 ? undefined : snapshot.document.workbook.sheets[sheetIndex];
  if (sheet === null || documentSheet === undefined) return null;
  const resolver = createPresentationResolver({
    document: snapshot.document,
    formulaValues: new Map(
      snapshot.calculation.values.map(({ address, value }) => [address, value]),
    ),
    cache: createPresentationCache({
      maximumEntries: 100_000,
      maximumBytes: 32 * 1024 * 1024,
    }),
    revisions: {
      document: snapshot.revision,
      calculation: snapshot.calculation.revision,
      condition: snapshot.revision,
      style: snapshot.revision,
      environment: snapshot.revision,
    },
    environment: {
      locale: snapshot.document.workbook.settings.localeHint ?? 'en-US',
      timeZone: 'UTC',
      dateSystem: snapshot.document.workbook.settings.dateSystem,
      target: 'print',
    },
  });
  const presentations = {
    resolve: (point: { readonly row: number; readonly column: number }) =>
      resolver.resolve({ sheetId: documentSheet.id, ...point }),
  };
  const layout = createPrintLayout(sheet, {
    paperSize: options.paper,
    orientation: options.orientation,
    defaultStyle,
    presentations,
  });
  const displayList = createPrintDisplayList({
    pages: layout.pages.map((page) => ({
      width: layout.paper.width,
      height: layout.paper.height,
      cells: page.cells.map((cell) => ({
        rect: {
          x: layout.contentLeft + cell.rect.left * layout.scale,
          y: layout.paper.padding + (cell.rect.top - page.contentTop) * layout.scale,
          width: cell.rect.width * layout.scale,
          height: cell.rect.height * layout.scale,
        },
        presentation: presentations.resolve(cell),
      })),
    })),
    fontMetrics:
      options.fontMetrics ??
      createFontMetrics({
        fonts: { Arial: { averageAdvance: 7, lineHeight: 13 } },
        fallbackFont: 'Arial',
        fallback: { averageAdvance: 7, lineHeight: 13 },
      }),
  });
  if (displayList.diagnostics.some(({ severity }) => severity === 'error')) {
    throw new Error(
      `Print display list failed: ${displayList.diagnostics
        .filter(({ severity }) => severity === 'error')
        .map(({ code }) => code)
        .join(', ')}`,
    );
  }
  return mountPrintDisplayPages(displayList, options);
}
