import type { SpreadsheetDocument } from '../../document';
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
              row: table.range.start.row + 1,
              column: table.range.start.column + columnIndex,
            },
            end: {
              row: table.range.end.row,
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
