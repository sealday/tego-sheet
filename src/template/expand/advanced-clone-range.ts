import type { DocumentCellRange, Sheet } from '../../document';
import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type { TemplateIRBinding } from '../model';
import { translatedCell, valueInput } from './advanced-internals';

export function cloneRange(
  sheet: Sheet,
  range: DocumentCellRange,
  rowCopies: number,
  columnCopies: number,
  items: readonly unknown[],
  axis: 'vertical' | 'horizontal' | 'both',
  valueBindings: readonly Extract<TemplateIRBinding, { readonly type: 'value' }>[],
  data: unknown,
  formatters: TemplateFormatterRegistry,
): { readonly sheet: Sheet; readonly mappings: readonly DocumentCellRange[] } {
  const height = range.end.row - range.start.row + 1;
  const width = range.end.column - range.start.column + 1;
  const rowDeltaTotal = height * (rowCopies - 1);
  const columnDeltaTotal = width * (columnCopies - 1);
  const sourceCells = sheet.cells.filter(
    ({ row, column }) =>
      row >= range.start.row &&
      row <= range.end.row &&
      column >= range.start.column &&
      column <= range.end.column,
  );
  const cellsOutside = sheet.cells
    .filter(
      ({ row, column }) =>
        row < range.start.row ||
        row > range.end.row ||
        column < range.start.column ||
        column > range.end.column,
    )
    .map((entry) => ({
      ...entry,
      row: entry.row > range.end.row ? entry.row + rowDeltaTotal : entry.row,
      column: entry.column > range.end.column ? entry.column + columnDeltaTotal : entry.column,
    }));
  const copies = [];
  const mappings: DocumentCellRange[] = [];
  for (let rowIndex = 0; rowIndex < rowCopies; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCopies; columnIndex += 1) {
      const rowDelta = rowIndex * height;
      const columnDelta = columnIndex * width;
      mappings.push({
        sheetId: range.sheetId,
        start: { row: range.start.row + rowDelta, column: range.start.column + columnDelta },
        end: { row: range.end.row + rowDelta, column: range.end.column + columnDelta },
      });
      const item =
        axis === 'vertical'
          ? items[rowIndex]
          : axis === 'horizontal'
            ? items[columnIndex]
            : Array.isArray(items[rowIndex])
              ? (items[rowIndex] as readonly unknown[])[columnIndex]
              : undefined;
      const itemIndex = axis === 'horizontal' ? columnIndex : rowIndex;
      const scope = {
        root: data,
        item,
        index: itemIndex,
        first: itemIndex === 0,
        last: itemIndex === (axis === 'horizontal' ? columnCopies - 1 : rowCopies - 1),
      };
      for (const entry of sourceCells) {
        const valueBinding = valueBindings.find(
          ({ target }) => target.row === entry.row && target.column === entry.column,
        );
        let cell = translatedCell(entry.cell, rowDelta, columnDelta);
        if (valueBinding !== undefined) {
          let value = evaluateTemplateExpression(valueBinding.expression, scope, formatters);
          if (valueBinding.formatter !== undefined) {
            const formatter = Object.getOwnPropertyDescriptor(
              formatters,
              valueBinding.formatter,
            )?.value;
            if (typeof formatter === 'function') value = formatter(value);
          }
          cell = { ...cell, input: valueInput(value) };
        }
        copies.push({
          ...entry,
          row: entry.row + rowDelta,
          column: entry.column + columnDelta,
          cell,
        });
      }
      for (const valueBinding of valueBindings.filter(
        ({ target }) =>
          !sourceCells.some(({ row, column }) => row === target.row && column === target.column),
      )) {
        copies.push({
          row: valueBinding.target.row + rowDelta,
          column: valueBinding.target.column + columnDelta,
          cell: {
            input: valueInput(
              evaluateTemplateExpression(valueBinding.expression, scope, formatters),
            ),
          },
        });
      }
    }
  }
  const sourceRows = sheet.rows.filter(
    ({ index }) => index >= range.start.row && index <= range.end.row,
  );
  const rows = [
    ...sheet.rows.filter(({ index }) => index < range.start.row),
    ...Array.from({ length: rowCopies }, (_, copy) =>
      sourceRows.map((row) => ({ ...row, index: row.index + copy * height })),
    ).flat(),
    ...sheet.rows
      .filter(({ index }) => index > range.end.row)
      .map((row) => ({ ...row, index: row.index + rowDeltaTotal })),
  ];
  const sourceColumns = sheet.columns.filter(
    ({ index }) => index >= range.start.column && index <= range.end.column,
  );
  const columns = [
    ...sheet.columns.filter(({ index }) => index < range.start.column),
    ...Array.from({ length: columnCopies }, (_, copy) =>
      sourceColumns.map((column) => ({ ...column, index: column.index + copy * width })),
    ).flat(),
    ...sheet.columns
      .filter(({ index }) => index > range.end.column)
      .map((column) => ({ ...column, index: column.index + columnDeltaTotal })),
  ];
  const sourceMerges = sheet.merges.filter(
    ({ start, end }) =>
      start.row >= range.start.row &&
      end.row <= range.end.row &&
      start.column >= range.start.column &&
      end.column <= range.end.column,
  );
  const merges = [
    ...sheet.merges
      .filter((merge) => !sourceMerges.includes(merge))
      .map((merge) => ({
        start: {
          row: merge.start.row > range.end.row ? merge.start.row + rowDeltaTotal : merge.start.row,
          column:
            merge.start.column > range.end.column
              ? merge.start.column + columnDeltaTotal
              : merge.start.column,
        },
        end: {
          row: merge.end.row > range.end.row ? merge.end.row + rowDeltaTotal : merge.end.row,
          column:
            merge.end.column > range.end.column
              ? merge.end.column + columnDeltaTotal
              : merge.end.column,
        },
      })),
    ...Array.from({ length: rowCopies }, (_, rowIndex) =>
      Array.from({ length: columnCopies }, (_, columnIndex) =>
        sourceMerges.map((merge) => ({
          start: {
            row: merge.start.row + rowIndex * height,
            column: merge.start.column + columnIndex * width,
          },
          end: {
            row: merge.end.row + rowIndex * height,
            column: merge.end.column + columnIndex * width,
          },
        })),
      ).flat(),
    ).flat(),
  ];
  return {
    sheet: {
      ...sheet,
      cells: [...cellsOutside, ...copies].sort(
        (left, right) => left.row - right.row || left.column - right.column,
      ),
      rows,
      columns,
      merges,
      ...(sheet.rowCount === undefined ? {} : { rowCount: sheet.rowCount + rowDeltaTotal }),
      ...(sheet.columnCount === undefined
        ? {}
        : { columnCount: sheet.columnCount + columnDeltaTotal }),
    },
    mappings,
  };
}
