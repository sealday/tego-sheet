import type { DocumentCellRange, SpreadsheetDocument } from '../../document';
import type { SpreadsheetTemplate, TemplateIRBinding } from '../model';
import { freeze, translatedCell } from './advanced-internals';

interface ExpandedChild {
  readonly document?: SpreadsheetDocument;
  readonly structuralMappings: readonly {
    readonly source: DocumentCellRange;
    readonly generated: DocumentCellRange;
  }[];
}

export function materializeSubtemplate(
  document: SpreadsheetDocument,
  binding: Extract<TemplateIRBinding, { readonly type: 'subtemplate' }>,
  registered: SpreadsheetTemplate,
  childExpansion: ExpandedChild,
):
  | {
      readonly document: SpreadsheetDocument;
      readonly generatedRange: DocumentCellRange;
    }
  | undefined {
  if (childExpansion.document === undefined) return undefined;
  const sheetIndex = document.workbook.sheets.findIndex(({ id }) => id === binding.range.sheetId);
  if (sheetIndex < 0) return { document, generatedRange: binding.range };
  const sheet = document.workbook.sheets[sheetIndex]!;
  const childCoordinates = registered.bindings.flatMap((candidate) =>
    candidate.type === 'value'
      ? [
          {
            sheetId: candidate.target.sheetId,
            start: { row: candidate.target.row, column: candidate.target.column },
            end: { row: candidate.target.row, column: candidate.target.column },
          },
        ]
      : [candidate.range],
  );
  const sourceSheetId = childCoordinates[0]?.sheetId ?? binding.range.sheetId;
  const sourceSheet = childExpansion.document.workbook.sheets.find(
    ({ id }) => id === sourceSheetId,
  );
  if (sourceSheet === undefined) return undefined;
  const sourceRange = childCoordinates.reduce<DocumentCellRange>(
    (range, candidate) => ({
      sheetId: sourceSheetId,
      start: {
        row: Math.min(range.start.row, candidate.start.row),
        column: Math.min(range.start.column, candidate.start.column),
      },
      end: {
        row: Math.max(range.end.row, candidate.end.row),
        column: Math.max(range.end.column, candidate.end.column),
      },
    }),
    childCoordinates[0] ?? {
      sheetId: sourceSheetId,
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 },
    },
  );
  const expandedRange = childExpansion.structuralMappings.reduce(
    (range, mapping) =>
      mapping.source.sheetId === sourceSheetId
        ? {
            ...range,
            end: {
              row: Math.max(range.end.row, mapping.generated.end.row),
              column: Math.max(range.end.column, mapping.generated.end.column),
            },
          }
        : range,
    sourceRange,
  );
  const sourceHeight = expandedRange.end.row - expandedRange.start.row + 1;
  const sourceWidth = expandedRange.end.column - expandedRange.start.column + 1;
  const targetHeight = binding.range.end.row - binding.range.start.row + 1;
  const targetWidth = binding.range.end.column - binding.range.start.column + 1;
  const rowDelta = sourceHeight - targetHeight;
  const columnDelta = sourceWidth - targetWidth;
  const pastedCells = sourceSheet.cells
    .filter(
      ({ row, column }) =>
        row >= expandedRange.start.row &&
        row <= expandedRange.end.row &&
        column >= expandedRange.start.column &&
        column <= expandedRange.end.column,
    )
    .map((entry) => {
      const row = binding.range.start.row + entry.row - expandedRange.start.row;
      const column = binding.range.start.column + entry.column - expandedRange.start.column;
      return {
        ...entry,
        row,
        column,
        cell: translatedCell(entry.cell, row - entry.row, column - entry.column),
      };
    });
  const cells = [
    ...sheet.cells
      .filter(
        ({ row, column }) =>
          row < binding.range.start.row ||
          row > binding.range.end.row ||
          column < binding.range.start.column ||
          column > binding.range.end.column,
      )
      .map((entry) => ({
        ...entry,
        row: entry.row > binding.range.end.row ? entry.row + rowDelta : entry.row,
        column: entry.column > binding.range.end.column ? entry.column + columnDelta : entry.column,
      })),
    ...pastedCells,
  ].sort((left, right) => left.row - right.row || left.column - right.column);
  const pastedMerges = sourceSheet.merges
    .filter(
      ({ start, end }) =>
        start.row >= expandedRange.start.row &&
        end.row <= expandedRange.end.row &&
        start.column >= expandedRange.start.column &&
        end.column <= expandedRange.end.column,
    )
    .map((merge) => ({
      start: {
        row: binding.range.start.row + merge.start.row - expandedRange.start.row,
        column: binding.range.start.column + merge.start.column - expandedRange.start.column,
      },
      end: {
        row: binding.range.start.row + merge.end.row - expandedRange.start.row,
        column: binding.range.start.column + merge.end.column - expandedRange.start.column,
      },
    }));
  return freeze({
    document: {
      ...document,
      workbook: {
        ...document.workbook,
        sheets: document.workbook.sheets.map((candidate, index) =>
          index === sheetIndex
            ? {
                ...candidate,
                cells,
                merges: [
                  ...candidate.merges.filter(
                    ({ start, end }) =>
                      end.row < binding.range.start.row ||
                      start.row > binding.range.end.row ||
                      end.column < binding.range.start.column ||
                      start.column > binding.range.end.column,
                  ),
                  ...pastedMerges,
                ],
              }
            : candidate,
        ),
      },
    },
    generatedRange: {
      sheetId: binding.range.sheetId,
      start: binding.range.start,
      end: {
        row: binding.range.start.row + sourceHeight - 1,
        column: binding.range.start.column + sourceWidth - 1,
      },
    },
  });
}
