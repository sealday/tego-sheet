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

/** Immutable row-order projection produced by table-local filters and sorting. */
export interface StructuredTableViewResult {
  readonly sourceRevision: string;
  readonly rowIndices: readonly number[];
}

/** Bounded scalar reader used to execute a table view against one immutable revision. */
export interface StructuredTableValueSource {
  readonly revision: string;
  readonly read: (row: number, column: number) => unknown;
}

function compareTableValues(left: unknown, right: unknown): number {
  if (typeof left === 'number' && Number.isFinite(left)) {
    return typeof right === 'number' && Number.isFinite(right) ? left - right : -1;
  }
  if (typeof right === 'number' && Number.isFinite(right)) return 1;
  return String(left ?? '').localeCompare(String(right ?? ''), 'en-US');
}

/** Executes persistent table filter/sort metadata without mutating worksheet rows. */
export function executeStructuredTableView(
  table: StructuredTable,
  source: StructuredTableValueSource,
  options: { readonly maximumRows?: number; readonly signal?: AbortSignal } = {},
): StructuredTableViewResult {
  const filterColumns = [
    ...(table.filter?.filters.map(({ column }) => column) ?? []),
    ...(table.filter?.sort === undefined || table.filter.sort === null
      ? []
      : [table.filter.sort.column]),
  ];
  if (
    filterColumns.some(
      (column) =>
        !Number.isSafeInteger(column) ||
        column < table.range.start.column ||
        column > table.range.end.column,
    )
  ) {
    throw new RangeError('Structured table filter or sort column must be within the table range');
  }
  const maximumRows = options.maximumRows ?? 100_000;
  if (!Number.isSafeInteger(maximumRows) || maximumRows <= 0) {
    throw new RangeError('Structured table row limit must be a positive safe integer');
  }
  const start = table.range.start.row + (table.headerRows ?? 1);
  const end = table.range.end.row - (table.totalsRow === true ? 1 : 0);
  const count = Math.max(0, end - start + 1);
  if (count > maximumRows) throw new RangeError('Structured table row limit exceeded');
  const filters = table.filter?.filters ?? [];
  const rows: { row: number; ordinal: number }[] = [];
  for (let row = start; row <= end; row += 1) {
    if (options.signal?.aborted === true)
      throw new DOMException('Table view aborted', 'AbortError');
    const visible = filters.every((filter) => {
      if (filter.operator === 'all') return true;
      const value = String(source.read(row, filter.column) ?? '');
      return filter.values.includes(value);
    });
    if (visible) rows.push({ row, ordinal: rows.length });
  }
  const sort = table.filter?.sort;
  if (sort !== undefined && sort !== null) {
    rows.sort((left, right) => {
      const order = compareTableValues(
        source.read(left.row, sort.column),
        source.read(right.row, sort.column),
      );
      return (sort.direction === 'asc' ? order : -order) || left.ordinal - right.ordinal;
    });
  }
  return Object.freeze({
    sourceRevision: source.revision,
    rowIndices: Object.freeze(rows.map(({ row }) => row)),
  });
}
