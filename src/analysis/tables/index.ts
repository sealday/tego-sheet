import type { SpreadsheetDocument } from '../../document';
import type { DocumentCellAddress, DocumentCellRange, StructuredTable } from '../../document';
import type {
  FormulaTableBindingRequest,
  FormulaTableBindingResolver,
  FormulaTableBindingResult,
} from '../../formula';

/** Supplies the current immutable document to a persistent table resolver. */
export type StructuredTableDocumentProvider = () => SpreadsheetDocument;

/**
 * Resolves persisted table display names to stable identities and current data-body ranges.
 *
 * A provider can be supplied by controllers whose immutable document snapshot changes over time.
 */
export function createStructuredTableResolver(
  source: SpreadsheetDocument | StructuredTableDocumentProvider,
): FormulaTableBindingResolver {
  const document = typeof source === 'function' ? source : () => source;
  return Object.freeze({
    resolve(request: FormulaTableBindingRequest): FormulaTableBindingResult {
      const snapshot = document();
      const tableKey = request.tableName.toLocaleLowerCase('en-US');
      const columnKey = request.columnName.toLocaleLowerCase('en-US');
      for (const sheet of snapshot.workbook.sheets) {
        const table = sheet.tables.find(({ name }) => name.toLocaleLowerCase('en-US') === tableKey);
        if (table === undefined) continue;
        const columnIndex = table.columns.findIndex(
          ({ name }) => name.toLocaleLowerCase('en-US') === columnKey,
        );
        if (columnIndex < 0) {
          return {
            status: 'invalid',
            message: `Unknown structured reference ${request.tableName}[${request.columnName}]`,
          };
        }
        return {
          status: 'resolved',
          tableId: table.id,
          columnId: table.columns[columnIndex]!.id,
          range: {
            sheetId: table.range.sheetId,
            start: {
              row: table.range.start.row + (table.headerRows ?? 1),
              column: table.range.start.column + columnIndex,
            },
            end: {
              row: table.range.end.row - (table.totalsRow === true ? 1 : 0),
              column: table.range.start.column + columnIndex,
            },
          },
        };
      }
      return {
        status: 'invalid',
        message: `Unknown structured reference ${request.tableName}[${request.columnName}]`,
      };
    },
  });
}

/** Result of deciding whether a direct append may expand a structured table. */
export type StructuredTableAutoExpandResult =
  | { readonly status: 'unchanged'; readonly table: StructuredTable }
  | { readonly status: 'expanded'; readonly table: StructuredTable }
  | {
      readonly status: 'rejected';
      readonly code:
        | 'TABLE_AUTO_EXPAND_DISABLED'
        | 'TABLE_AUTO_EXPAND_OUTSIDE_BOUNDARY'
        | 'TABLE_RANGE_OVERLAP'
        | 'TABLE_CELL_LIMIT_EXCEEDED';
    };

function overlaps(left: DocumentCellRange, right: DocumentCellRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.start.row <= right.end.row &&
    right.start.row <= left.end.row &&
    left.start.column <= right.end.column &&
    right.start.column <= left.end.column
  );
}

/** Plans a bounded one-row auto expansion without mutating the source table. */
export function planStructuredTableAutoExpand(
  table: StructuredTable,
  append: DocumentCellAddress,
  occupiedRanges: readonly DocumentCellRange[] = [],
  maximumCells = 1_000_000,
): StructuredTableAutoExpandResult {
  if (table.autoExpand !== true) {
    return { status: 'rejected', code: 'TABLE_AUTO_EXPAND_DISABLED' };
  }
  if (
    append.sheetId !== table.range.sheetId ||
    append.row !== table.range.end.row + 1 ||
    append.column < table.range.start.column ||
    append.column > table.range.end.column
  ) {
    return { status: 'rejected', code: 'TABLE_AUTO_EXPAND_OUTSIDE_BOUNDARY' };
  }
  const range = {
    ...table.range,
    end: { row: append.row, column: table.range.end.column },
  };
  const cellCount =
    (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1);
  if (!Number.isSafeInteger(maximumCells) || maximumCells <= 0 || cellCount > maximumCells) {
    return { status: 'rejected', code: 'TABLE_CELL_LIMIT_EXCEEDED' };
  }
  if (occupiedRanges.some((occupied) => overlaps(range, occupied))) {
    return { status: 'rejected', code: 'TABLE_RANGE_OVERLAP' };
  }
  return { status: 'expanded', table: Object.freeze({ ...table, range: Object.freeze(range) }) };
}
