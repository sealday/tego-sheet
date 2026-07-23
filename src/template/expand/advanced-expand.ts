import type {
  Cell,
  Diagnostic,
  DocumentCellRange,
  Sheet,
  SpreadsheetDocument,
} from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';
import { compileSpreadsheetTemplate } from '../compiler';
import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type {
  CompiledTemplate,
  RenderLimits,
  TemplateIRBinding,
  TemplateRegionNode,
} from '../model';
import { expandTemplate } from './expand';

/** One source-to-generated structural coordinate mapping. */
export interface StructuralMapping {
  /** Binding that generated the mapping. */
  readonly bindingId: string;
  /** Zero-based collection item index. */
  readonly itemIndex: number;
  /** Original source region. */
  readonly source: DocumentCellRange;
  /** Generated output region. */
  readonly generated: DocumentCellRange;
  /** Applied floating-object policy. */
  readonly objectPolicy?: 'per-item' | 'shared' | 'forbidden';
}

/** Atomic TP2 expansion result. */
export interface AdvancedExpansionResult {
  /** Complete expanded snapshot, absent on any error. */
  readonly document?: SpreadsheetDocument;
  /** Ordered expansion diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Deterministic structural coordinate mappings. */
  readonly structuralMappings: readonly StructuralMapping[];
  /** Hard page boundaries keyed by generated sheet. */
  readonly forcedPageBreaks: ReadonlyMap<string, readonly number[]>;
}

interface Scope {
  readonly root: unknown;
  readonly item?: unknown;
  readonly parent?: unknown;
  readonly index?: number;
  readonly first?: boolean;
  readonly last?: boolean;
}

function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (value !== null && typeof value === 'object' && !(value instanceof Map)) {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

function error(code: string, message: string, bindingId?: string): Diagnostic {
  return freeze({
    code,
    severity: 'error',
    domain: 'template',
    stage: 'expand',
    message,
    ...(bindingId === undefined ? {} : { location: { bindingId: bindingId as never } }),
  });
}

function translatedCell(cell: Cell, rowDelta: number, columnDelta: number): Cell {
  if (cell.input.type !== 'formula' || (rowDelta === 0 && columnDelta === 0)) return cell;
  try {
    return {
      ...cell,
      input: {
        type: 'formula',
        source: renderFormula(
          translateFormula(parseFormula(cell.input.source), { rowDelta, columnDelta }),
        ),
      },
    };
  } catch {
    return cell;
  }
}

function cloneRange(
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

function collection(
  binding: Extract<
    TemplateIRBinding,
    {
      readonly type:
        | 'repeat-columns'
        | 'repeat-range'
        | 'repeat-page'
        | 'repeat-sheet'
        | 'subtemplate';
    }
  >,
  scope: Scope,
  formatters: TemplateFormatterRegistry,
): readonly unknown[] {
  const value = evaluateTemplateExpression(binding.source, scope, formatters);
  return Array.isArray(value) ? value : [];
}

function safeSheetName(value: unknown, fallback: string): string {
  const sanitized = String(value ?? fallback)
    .replace(/[:\\/?*[\]]/gu, '_')
    .trim()
    .slice(0, 31);
  return sanitized || fallback;
}

function valueInput(value: unknown): Cell['input'] {
  if (value === undefined || value === null) return { type: 'blank' };
  if (typeof value === 'string') return { type: 'string', value };
  if (typeof value === 'number' && Number.isFinite(value)) return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'string', value: JSON.stringify(value) };
}

function expandNestedRows(
  document: SpreadsheetDocument,
  compiled: CompiledTemplate,
  data: unknown,
  formatters: TemplateFormatterRegistry,
  limits: RenderLimits,
  signal?: AbortSignal,
): {
  readonly document?: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly mappings: readonly StructuralMapping[];
  readonly breaks: ReadonlyMap<string, readonly number[]>;
} {
  const byId = new Map(compiled.ir.bindings.map((binding) => [binding.id, binding]));
  const isMixedVerticalTree = (node: TemplateRegionNode, root = false): boolean => {
    const type = byId.get(node.bindingId)?.type;
    return (
      (root ? type === 'repeat-rows' : type === 'repeat-rows' || type === 'repeat-columns') &&
      node.children.every((child) => isMixedVerticalTree(child))
    );
  };
  const roots = (compiled.ir.regionTree ?? []).filter((node) => isMixedVerticalTree(node, true));
  if (roots.length === 0) {
    return { document, diagnostics: [], mappings: [], breaks: new Map() };
  }
  const valueBindings = compiled.ir.bindings.filter(
    (binding): binding is Extract<TemplateIRBinding, { readonly type: 'value' }> =>
      binding.type === 'value',
  );
  const diagnostics: Diagnostic[] = [];
  const mappings: StructuralMapping[] = [];
  const breaks = new Map<string, number[]>();
  const sheets = [...document.workbook.sheets];
  let totalCells = sheets.reduce((count, sheet) => count + sheet.cells.length, 0);
  let totalRows = 0;

  for (const root of [...roots].sort(
    (left, right) => right.range.start.row - left.range.start.row,
  )) {
    const sheetIndex = sheets.findIndex(({ id }) => id === root.range.sheetId);
    if (sheetIndex < 0) continue;
    const sourceSheet = sheets[sheetIndex]!;

    interface Fragment {
      readonly cells: Sheet['cells'][number][];
      readonly rows: Sheet['rows'][number][];
      readonly merges: Sheet['merges'];
      readonly height: number;
    }
    interface HorizontalFragment extends Fragment {
      readonly width: number;
    }

    const renderNode = (
      node: TemplateRegionNode,
      parentScope: Scope,
      destinationStart: number,
    ): Fragment => {
      const binding = byId.get(node.bindingId) as
        | Extract<TemplateIRBinding, { readonly type: 'repeat-rows' }>
        | undefined;
      if (binding === undefined) return { cells: [], rows: [], merges: [], height: 0 };
      const resolved = evaluateTemplateExpression(binding.source, parentScope, formatters);
      const items = Array.isArray(resolved) ? resolved : [];
      const copies =
        items.length === 0 && binding.empty === 'keep-template-row' ? [undefined] : items;
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      let height = 0;
      for (const [index, item] of copies.entries()) {
        if (signal?.aborted) return { cells: [], rows: [], merges: [], height: 0 };
        const fragment = renderRange(
          node.range,
          node.children,
          {
            root: data,
            item,
            parent: parentScope.item,
            index,
            first: index === 0,
            last: index === copies.length - 1,
          },
          destinationStart + height,
        );
        mappings.push({
          bindingId: binding.id,
          itemIndex: index,
          source: binding.range,
          generated: {
            sheetId: binding.range.sheetId,
            start: { ...binding.range.start, row: destinationStart + height },
            end: {
              ...binding.range.end,
              row: destinationStart + height + Math.max(0, fragment.height - 1),
            },
          },
        });
        if (binding.pageBreak === 'before-each-item' && index > 0) {
          const values = breaks.get(binding.range.sheetId) ?? [];
          values.push(destinationStart + height);
          breaks.set(binding.range.sheetId, values);
        }
        cells.push(...fragment.cells);
        rows.push(...fragment.rows);
        merges.push(...fragment.merges);
        height += fragment.height;
      }
      return { cells, rows, merges, height };
    };

    const renderRange = (
      range: DocumentCellRange,
      children: readonly TemplateRegionNode[],
      scope: Scope,
      destinationStart: number,
    ): Fragment => {
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      const rowMap = new Map<number, number>();
      const sortedChildren = [...children].sort(
        (left, right) => left.range.start.row - right.range.start.row,
      );
      let childIndex = 0;
      let sourceRow = range.start.row;
      let height = 0;
      while (sourceRow <= range.end.row) {
        const child = sortedChildren[childIndex];
        if (child !== undefined && child.range.start.row === sourceRow) {
          const childType = byId.get(child.bindingId)?.type;
          const fragment =
            childType === 'repeat-columns'
              ? renderHorizontalNode(
                  child,
                  scope,
                  destinationStart + height,
                  child.range.start.column,
                )
              : renderNode(child, scope, destinationStart + height);
          cells.push(...fragment.cells);
          rows.push(...fragment.rows);
          merges.push(...fragment.merges);
          height += fragment.height;
          sourceRow = child.range.end.row + 1;
          childIndex += 1;
          continue;
        }
        const generatedRow = destinationStart + height;
        rowMap.set(sourceRow, generatedRow);
        const rowDelta = generatedRow - sourceRow;
        const sourceCells = sourceSheet.cells.filter(
          ({ row, column }) =>
            row === sourceRow && column >= range.start.column && column <= range.end.column,
        );
        const targets = valueBindings.filter(
          ({ target }) =>
            target.sheetId === sourceSheet.id &&
            target.row === sourceRow &&
            target.column >= range.start.column &&
            target.column <= range.end.column,
        );
        for (const entry of sourceCells) {
          const valueBinding = targets.find(({ target }) => target.column === entry.column);
          let cell = translatedCell(entry.cell, rowDelta, 0);
          if (valueBinding !== undefined) {
            let value = evaluateTemplateExpression(valueBinding.expression, scope, formatters);
            if (value === undefined) {
              diagnostics.push(
                error(
                  'MISSING_DATA',
                  `Binding ${valueBinding.id} resolved to a missing value`,
                  valueBinding.id,
                ),
              );
              value = null;
            }
            cell = { ...cell, input: valueInput(value) };
          }
          cells.push({ ...entry, row: generatedRow, cell });
        }
        for (const valueBinding of targets.filter(
          ({ target }) => !sourceCells.some(({ column }) => column === target.column),
        )) {
          cells.push({
            row: generatedRow,
            column: valueBinding.target.column,
            cell: {
              input: valueInput(
                evaluateTemplateExpression(valueBinding.expression, scope, formatters),
              ),
            },
          });
        }
        const row = sourceSheet.rows.find(({ index }) => index === sourceRow);
        if (row !== undefined) rows.push({ ...row, index: generatedRow });
        sourceRow += 1;
        height += 1;
      }
      for (const merge of sourceSheet.merges.filter(
        ({ start, end }) =>
          start.row >= range.start.row &&
          end.row <= range.end.row &&
          start.column >= range.start.column &&
          end.column <= range.end.column,
      )) {
        const startRow = rowMap.get(merge.start.row);
        const endRow = rowMap.get(merge.end.row);
        if (startRow !== undefined && endRow !== undefined) {
          merges.push({
            start: { ...merge.start, row: startRow },
            end: { ...merge.end, row: endRow },
          });
        }
      }
      return { cells, rows, merges, height };
    };

    const renderHorizontalNode = (
      node: TemplateRegionNode,
      parentScope: Scope,
      destinationRow: number,
      destinationColumn: number,
    ): HorizontalFragment => {
      const binding = byId.get(node.bindingId);
      if (binding?.type !== 'repeat-columns') {
        return { cells: [], rows: [], merges: [], height: 0, width: 0 };
      }
      const resolved = evaluateTemplateExpression(binding.source, parentScope, formatters);
      const items = Array.isArray(resolved) ? resolved : [];
      const copies =
        items.length === 0 && binding.empty === 'keep-template-row' ? [undefined] : items;
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      let width = 0;
      let height = 0;
      for (const [index, item] of copies.entries()) {
        const fragment = renderHorizontalRange(
          node.range,
          node.children,
          {
            root: data,
            item,
            parent: parentScope.item,
            index,
            first: index === 0,
            last: index === copies.length - 1,
          },
          destinationRow,
          destinationColumn + width,
        );
        mappings.push({
          bindingId: binding.id,
          itemIndex: index,
          source: binding.range,
          generated: {
            sheetId: binding.range.sheetId,
            start: { row: destinationRow, column: destinationColumn + width },
            end: {
              row: destinationRow + Math.max(0, fragment.height - 1),
              column: destinationColumn + width + Math.max(0, fragment.width - 1),
            },
          },
        });
        cells.push(...fragment.cells);
        if (index === 0 || fragment.height > height) {
          rows.splice(0, rows.length, ...fragment.rows);
          height = fragment.height;
        }
        merges.push(...fragment.merges);
        width += fragment.width;
      }
      return { cells, rows, merges, height, width };
    };

    const renderHorizontalRange = (
      range: DocumentCellRange,
      children: readonly TemplateRegionNode[],
      scope: Scope,
      destinationRow: number,
      destinationColumn: number,
    ): HorizontalFragment => {
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      const rowMap = new Map<number, number>();
      const sortedChildren = [...children].sort(
        (left, right) => left.range.start.row - right.range.start.row,
      );
      let childIndex = 0;
      let sourceRow = range.start.row;
      let height = 0;
      while (sourceRow <= range.end.row) {
        const child = sortedChildren[childIndex];
        if (child !== undefined && child.range.start.row === sourceRow) {
          const childType = byId.get(child.bindingId)?.type;
          const childFragment =
            childType === 'repeat-columns'
              ? renderHorizontalNode(
                  child,
                  scope,
                  destinationRow + height,
                  destinationColumn + (child.range.start.column - range.start.column),
                )
              : renderNode(child, scope, destinationRow + height);
          const columnDelta = destinationColumn - range.start.column;
          cells.push(
            ...childFragment.cells.map((entry) => ({
              ...entry,
              column: entry.column + columnDelta,
              cell: translatedCell(entry.cell, 0, columnDelta),
            })),
          );
          rows.push(...childFragment.rows);
          merges.push(
            ...childFragment.merges.map((merge) => ({
              start: { ...merge.start, column: merge.start.column + columnDelta },
              end: { ...merge.end, column: merge.end.column + columnDelta },
            })),
          );
          height += childFragment.height;
          sourceRow = child.range.end.row + 1;
          childIndex += 1;
          continue;
        }
        const generatedRow = destinationRow + height;
        rowMap.set(sourceRow, generatedRow);
        const rowDelta = generatedRow - sourceRow;
        const columnDelta = destinationColumn - range.start.column;
        const sourceCells = sourceSheet.cells.filter(
          ({ row, column }) =>
            row === sourceRow && column >= range.start.column && column <= range.end.column,
        );
        const targets = valueBindings.filter(
          ({ target }) =>
            target.sheetId === sourceSheet.id &&
            target.row === sourceRow &&
            target.column >= range.start.column &&
            target.column <= range.end.column,
        );
        for (const entry of sourceCells) {
          const valueBinding = targets.find(({ target }) => target.column === entry.column);
          const cell =
            valueBinding === undefined
              ? translatedCell(entry.cell, rowDelta, columnDelta)
              : {
                  ...entry.cell,
                  input: valueInput(
                    evaluateTemplateExpression(valueBinding.expression, scope, formatters),
                  ),
                };
          cells.push({
            ...entry,
            row: generatedRow,
            column: entry.column + columnDelta,
            cell,
          });
        }
        for (const valueBinding of targets.filter(
          ({ target }) => !sourceCells.some(({ column }) => column === target.column),
        )) {
          cells.push({
            row: generatedRow,
            column: valueBinding.target.column + columnDelta,
            cell: {
              input: valueInput(
                evaluateTemplateExpression(valueBinding.expression, scope, formatters),
              ),
            },
          });
        }
        const row = sourceSheet.rows.find(({ index }) => index === sourceRow);
        if (row !== undefined) rows.push({ ...row, index: generatedRow });
        sourceRow += 1;
        height += 1;
      }
      const columnDelta = destinationColumn - range.start.column;
      for (const merge of sourceSheet.merges.filter(
        ({ start, end }) =>
          start.row >= range.start.row &&
          end.row <= range.end.row &&
          start.column >= range.start.column &&
          end.column <= range.end.column,
      )) {
        const startRow = rowMap.get(merge.start.row);
        const endRow = rowMap.get(merge.end.row);
        if (startRow !== undefined && endRow !== undefined) {
          merges.push({
            start: { row: startRow, column: merge.start.column + columnDelta },
            end: { row: endRow, column: merge.end.column + columnDelta },
          });
        }
      }
      return {
        cells,
        rows,
        merges,
        height,
        width: range.end.column - range.start.column + 1,
      };
    };

    const estimateNode = (
      node: TemplateRegionNode,
      parentScope: Scope,
    ): { readonly rows: number; readonly cells: number; readonly exceeded: boolean } => {
      const binding = byId.get(node.bindingId) as
        | Extract<TemplateIRBinding, { readonly type: 'repeat-rows' }>
        | undefined;
      if (binding === undefined) return { rows: 0, cells: 0, exceeded: false };
      const resolved = evaluateTemplateExpression(binding.source, parentScope, formatters);
      const items = Array.isArray(resolved) ? resolved : [];
      const copies =
        items.length === 0 && binding.empty === 'keep-template-row' ? [undefined] : items;
      let rows = 0;
      let cells = 0;
      for (const [index, item] of copies.entries()) {
        const scope: Scope = {
          root: data,
          item,
          parent: parentScope.item,
          index,
          first: index === 0,
          last: index === copies.length - 1,
        };
        const childRanges = node.children.map(({ range }) => range);
        const ownRows = Array.from(
          { length: node.range.end.row - node.range.start.row + 1 },
          (_, offset) => node.range.start.row + offset,
        ).filter(
          (row) => !childRanges.some((range) => row >= range.start.row && row <= range.end.row),
        );
        rows += ownRows.length;
        cells += ownRows.reduce((count, row) => {
          const sourceCount = sourceSheet.cells.filter((entry) => entry.row === row).length;
          const blankTargets = valueBindings.filter(
            ({ target }) =>
              target.sheetId === sourceSheet.id &&
              target.row === row &&
              !sourceSheet.cells.some(
                (entry) => entry.row === row && entry.column === target.column,
              ),
          ).length;
          return count + sourceCount + blankTargets;
        }, 0);
        for (const child of node.children) {
          const estimate = estimateNode(child, scope);
          rows += estimate.rows;
          cells += estimate.cells;
          if (
            estimate.exceeded ||
            totalRows + rows > limits.maxExpandedRows ||
            totalCells + cells > limits.maxExpandedCells
          ) {
            return { rows, cells, exceeded: true };
          }
        }
      }
      return {
        rows,
        cells,
        exceeded:
          totalRows + rows > limits.maxExpandedRows || totalCells + cells > limits.maxExpandedCells,
      };
    };

    const estimate = estimateNode(root, { root: data });
    if (estimate.exceeded) {
      return {
        diagnostics: [
          error(
            'EXPANSION_LIMIT_EXCEEDED',
            'Nested expansion exceeds configured limits',
            root.bindingId,
          ),
        ],
        mappings: [],
        breaks: new Map(),
      };
    }
    const fragment = renderNode(root, { root: data }, root.range.start.row);
    if (signal?.aborted) {
      return {
        diagnostics: [error('RENDER_ABORTED', 'Template rendering was aborted')],
        mappings: [],
        breaks: new Map(),
      };
    }
    const sourceHeight = root.range.end.row - root.range.start.row + 1;
    const delta = fragment.height - sourceHeight;
    const replaced = sourceSheet.cells.filter(
      ({ row }) => row >= root.range.start.row && row <= root.range.end.row,
    ).length;
    totalCells += fragment.cells.length - replaced;
    totalRows += fragment.height;
    if (
      totalCells > limits.maxExpandedCells ||
      totalRows > limits.maxExpandedRows ||
      diagnostics.some(({ severity }) => severity === 'error')
    ) {
      return {
        diagnostics:
          diagnostics.length > 0
            ? diagnostics
            : [
                error(
                  'EXPANSION_LIMIT_EXCEEDED',
                  'Nested expansion exceeds configured limits',
                  root.bindingId,
                ),
              ],
        mappings: [],
        breaks: new Map(),
      };
    }
    sheets[sheetIndex] = {
      ...sourceSheet,
      cells: [
        ...sourceSheet.cells.filter(({ row }) => row < root.range.start.row),
        ...fragment.cells,
        ...sourceSheet.cells
          .filter(({ row }) => row > root.range.end.row)
          .map((entry) => ({ ...entry, row: entry.row + delta })),
      ].sort((left, right) => left.row - right.row || left.column - right.column),
      rows: [
        ...sourceSheet.rows.filter(({ index }) => index < root.range.start.row),
        ...fragment.rows,
        ...sourceSheet.rows
          .filter(({ index }) => index > root.range.end.row)
          .map((row) => ({ ...row, index: row.index + delta })),
      ],
      merges: [
        ...sourceSheet.merges
          .filter(
            ({ start, end }) => end.row < root.range.start.row || start.row > root.range.end.row,
          )
          .map((merge) =>
            merge.start.row > root.range.end.row
              ? {
                  start: { ...merge.start, row: merge.start.row + delta },
                  end: { ...merge.end, row: merge.end.row + delta },
                }
              : merge,
          ),
        ...fragment.merges,
      ],
      ...(sourceSheet.rowCount === undefined ? {} : { rowCount: sourceSheet.rowCount + delta }),
    };
  }
  return {
    document: freeze({ ...document, workbook: { ...document.workbook, sheets } }),
    diagnostics,
    mappings,
    breaks,
  };
}

function expandNestedColumns(
  document: SpreadsheetDocument,
  compiled: CompiledTemplate,
  data: unknown,
  formatters: TemplateFormatterRegistry,
  limits: RenderLimits,
  signal?: AbortSignal,
): {
  readonly document?: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly mappings: readonly StructuralMapping[];
} {
  const byId = new Map(compiled.ir.bindings.map((binding) => [binding.id, binding]));
  const isHorizontalTree = (node: TemplateRegionNode): boolean =>
    byId.get(node.bindingId)?.type === 'repeat-columns' && node.children.every(isHorizontalTree);
  const roots = (compiled.ir.regionTree ?? []).filter(isHorizontalTree);
  if (roots.length === 0) return { document, diagnostics: [], mappings: [] };
  const values = compiled.ir.bindings.filter(
    (binding): binding is Extract<TemplateIRBinding, { readonly type: 'value' }> =>
      binding.type === 'value',
  );
  const mappings: StructuralMapping[] = [];
  const sheets = [...document.workbook.sheets];
  for (const root of [...roots].sort(
    (left, right) => right.range.start.column - left.range.start.column,
  )) {
    const sheetIndex = sheets.findIndex(({ id }) => id === root.range.sheetId);
    if (sheetIndex < 0) continue;
    const source = sheets[sheetIndex]!;
    interface Fragment {
      readonly cells: Sheet['cells'][number][];
      readonly columns: Sheet['columns'][number][];
      readonly merges: Sheet['merges'];
      readonly width: number;
    }
    const renderNode = (node: TemplateRegionNode, parent: Scope, destination: number): Fragment => {
      const binding = byId.get(node.bindingId) as
        | Extract<TemplateIRBinding, { readonly type: 'repeat-columns' }>
        | undefined;
      if (binding === undefined) return { cells: [], columns: [], merges: [], width: 0 };
      const resolved = evaluateTemplateExpression(binding.source, parent, formatters);
      const items = Array.isArray(resolved) ? resolved : [];
      const copies =
        items.length === 0 && binding.empty === 'keep-template-row' ? [undefined] : items;
      const cells: Sheet['cells'][number][] = [];
      const columns: Sheet['columns'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      let width = 0;
      for (const [index, item] of copies.entries()) {
        const fragment = renderRange(
          node.range,
          node.children,
          {
            root: data,
            item,
            parent: parent.item,
            index,
            first: index === 0,
            last: index === copies.length - 1,
          },
          destination + width,
        );
        mappings.push({
          bindingId: binding.id,
          itemIndex: index,
          source: binding.range,
          generated: {
            sheetId: binding.range.sheetId,
            start: { ...binding.range.start, column: destination + width },
            end: {
              ...binding.range.end,
              column: destination + width + Math.max(0, fragment.width - 1),
            },
          },
        });
        cells.push(...fragment.cells);
        columns.push(...fragment.columns);
        merges.push(...fragment.merges);
        width += fragment.width;
        if (
          width > (limits.maxExpandedColumns ?? 16_384) ||
          cells.length > limits.maxExpandedCells
        ) {
          return { cells: [], columns: [], merges: [], width: -1 };
        }
      }
      return { cells, columns, merges, width };
    };
    const renderRange = (
      range: DocumentCellRange,
      children: readonly TemplateRegionNode[],
      scope: Scope,
      destination: number,
    ): Fragment => {
      const cells: Sheet['cells'][number][] = [];
      const columns: Sheet['columns'][number][] = [];
      const merges: Sheet['merges'][number][] = [];
      const columnMap = new Map<number, number>();
      const sorted = [...children].sort(
        (left, right) => left.range.start.column - right.range.start.column,
      );
      let sourceColumn = range.start.column;
      let childIndex = 0;
      let width = 0;
      while (sourceColumn <= range.end.column) {
        const child = sorted[childIndex];
        if (child !== undefined && child.range.start.column === sourceColumn) {
          const fragment = renderNode(child, scope, destination + width);
          if (fragment.width < 0) return fragment;
          cells.push(...fragment.cells);
          columns.push(...fragment.columns);
          merges.push(...fragment.merges);
          width += fragment.width;
          sourceColumn = child.range.end.column + 1;
          childIndex += 1;
          continue;
        }
        const generatedColumn = destination + width;
        columnMap.set(sourceColumn, generatedColumn);
        const delta = generatedColumn - sourceColumn;
        const sourceCells = source.cells.filter(({ column }) => column === sourceColumn);
        const targets = values.filter(
          ({ target }) => target.sheetId === source.id && target.column === sourceColumn,
        );
        for (const entry of sourceCells) {
          const value = targets.find(({ target }) => target.row === entry.row);
          const cell =
            value === undefined
              ? translatedCell(entry.cell, 0, delta)
              : {
                  ...entry.cell,
                  input: valueInput(
                    evaluateTemplateExpression(value.expression, scope, formatters),
                  ),
                };
          cells.push({ ...entry, column: generatedColumn, cell });
        }
        for (const value of targets.filter(
          ({ target }) => !sourceCells.some(({ row }) => row === target.row),
        )) {
          cells.push({
            row: value.target.row,
            column: generatedColumn,
            cell: {
              input: valueInput(evaluateTemplateExpression(value.expression, scope, formatters)),
            },
          });
        }
        const column = source.columns.find(({ index }) => index === sourceColumn);
        if (column !== undefined) columns.push({ ...column, index: generatedColumn });
        sourceColumn += 1;
        width += 1;
      }
      for (const merge of source.merges.filter(
        ({ start, end }) =>
          start.column >= range.start.column &&
          end.column <= range.end.column &&
          start.row >= range.start.row &&
          end.row <= range.end.row,
      )) {
        const start = columnMap.get(merge.start.column);
        const end = columnMap.get(merge.end.column);
        if (start !== undefined && end !== undefined) {
          merges.push({
            start: { ...merge.start, column: start },
            end: { ...merge.end, column: end },
          });
        }
      }
      return { cells, columns, merges, width };
    };
    const fragment = renderNode(root, { root: data }, root.range.start.column);
    if (signal?.aborted) {
      return {
        diagnostics: [error('RENDER_ABORTED', 'Template rendering was aborted')],
        mappings: [],
      };
    }
    if (fragment.width < 0) {
      return {
        diagnostics: [
          error(
            'EXPANSION_LIMIT_EXCEEDED',
            'Nested horizontal expansion exceeds configured limits',
            root.bindingId,
          ),
        ],
        mappings: [],
      };
    }
    const sourceWidth = root.range.end.column - root.range.start.column + 1;
    const delta = fragment.width - sourceWidth;
    sheets[sheetIndex] = {
      ...source,
      cells: [
        ...source.cells.filter(({ column }) => column < root.range.start.column),
        ...fragment.cells,
        ...source.cells
          .filter(({ column }) => column > root.range.end.column)
          .map((entry) => ({ ...entry, column: entry.column + delta })),
      ].sort((left, right) => left.row - right.row || left.column - right.column),
      columns: [
        ...source.columns.filter(({ index }) => index < root.range.start.column),
        ...fragment.columns,
        ...source.columns
          .filter(({ index }) => index > root.range.end.column)
          .map((column) => ({ ...column, index: column.index + delta })),
      ],
      merges: [
        ...source.merges
          .filter(
            ({ start, end }) =>
              end.column < root.range.start.column || start.column > root.range.end.column,
          )
          .map((merge) =>
            merge.start.column > root.range.end.column
              ? {
                  start: { ...merge.start, column: merge.start.column + delta },
                  end: { ...merge.end, column: merge.end.column + delta },
                }
              : merge,
          ),
        ...fragment.merges,
      ],
    };
  }
  return {
    document: freeze({ ...document, workbook: { ...document.workbook, sheets } }),
    diagnostics: [],
    mappings,
  };
}

/** Expands TP2 structural bindings without publishing a partial document on failure. */
export function expandAdvancedTemplate(
  compiled: CompiledTemplate,
  data: unknown,
  limits: RenderLimits,
  formatters: TemplateFormatterRegistry = Object.freeze({}),
  signal?: AbortSignal,
): AdvancedExpansionResult {
  if (signal?.aborted) {
    return freeze({
      diagnostics: [error('RENDER_ABORTED', 'Template rendering was aborted')],
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  const bindingType = new Map(compiled.ir.bindings.map((binding) => [binding.id, binding.type]));
  const horizontalNestedIds = new Set<string>();
  const rowTreeIds = new Set<string>();
  const collectRowTree = (node: TemplateRegionNode): void => {
    rowTreeIds.add(node.bindingId);
    node.children.forEach(collectRowTree);
  };
  (compiled.ir.regionTree ?? [])
    .filter((node) => bindingType.get(node.bindingId) === 'repeat-rows')
    .forEach(collectRowTree);
  const collectHorizontal = (node: TemplateRegionNode): boolean => {
    if (
      bindingType.get(node.bindingId) !== 'repeat-columns' ||
      !node.children.every(collectHorizontal)
    ) {
      return false;
    }
    horizontalNestedIds.add(node.bindingId);
    node.children.forEach((child) => horizontalNestedIds.add(child.bindingId));
    return true;
  };
  (compiled.ir.regionTree ?? []).forEach(collectHorizontal);
  const advanced = compiled.ir.bindings.filter(
    (
      binding,
    ): binding is Extract<
      TemplateIRBinding,
      {
        readonly type:
          | 'repeat-columns'
          | 'repeat-range'
          | 'repeat-page'
          | 'repeat-sheet'
          | 'subtemplate';
      }
    > =>
      (binding.type === 'repeat-columns' ||
        binding.type === 'repeat-range' ||
        binding.type === 'repeat-page' ||
        binding.type === 'repeat-sheet' ||
        binding.type === 'subtemplate') &&
      !rowTreeIds.has(binding.id) &&
      !horizontalNestedIds.has(binding.id),
  );
  const tp1Bindings = compiled.ir.bindings.filter(
    (binding) =>
      binding.type === 'value' ||
      binding.type === 'repeat-rows' ||
      binding.type === 'conditional-range',
  );
  const nestedRanges = compiled.ir.bindings.flatMap((binding) =>
    binding.type === 'repeat-rows' ||
    binding.type === 'repeat-columns' ||
    binding.type === 'repeat-range' ||
    binding.type === 'repeat-page' ||
    binding.type === 'repeat-sheet' ||
    binding.type === 'subtemplate'
      ? [binding.range]
      : [],
  );
  const handledByNested = (binding: TemplateIRBinding): boolean =>
    nestedRanges.some((range) =>
      'target' in binding
        ? binding.target.sheetId === range.sheetId &&
          binding.target.row >= range.start.row &&
          binding.target.row <= range.end.row
        : binding.range.sheetId === range.sheetId &&
          binding.range.start.row >= range.start.row &&
          binding.range.end.row <= range.end.row,
    );
  const base = expandTemplate(
    compiled.sourceDocument,
    compiled.ir.template,
    tp1Bindings.filter((binding) => !handledByNested(binding)),
    data,
    formatters,
    limits,
    signal,
  );
  if (base.document === undefined) {
    return freeze({
      diagnostics: base.diagnostics,
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  const nested = expandNestedRows(base.document, compiled, data, formatters, limits, signal);
  if (nested.document === undefined) {
    return freeze({
      diagnostics: nested.diagnostics,
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  const nestedColumns = expandNestedColumns(
    nested.document,
    compiled,
    data,
    formatters,
    limits,
    signal,
  );
  if (nestedColumns.document === undefined) {
    return freeze({
      diagnostics: nestedColumns.diagnostics,
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  const estimates = advanced.map((binding) => ({
    binding,
    items: collection(binding, { root: data }, formatters),
  }));
  const sourceCells = compiled.sourceDocument.workbook.sheets.reduce(
    (count, sheet) => count + sheet.cells.length,
    0,
  );
  let estimatedRows = 0;
  let estimatedColumns = 0;
  const estimatedCells = estimates.reduce((count, { binding, items }) => {
    const source =
      compiled.sourceDocument.workbook.sheets
        .find(({ id }) => id === binding.range.sheetId)
        ?.cells.filter(
          ({ row, column }) =>
            row >= binding.range.start.row &&
            row <= binding.range.end.row &&
            column >= binding.range.start.column &&
            column <= binding.range.end.column,
        ).length ?? 0;
    const matrixRows = items.length;
    const matrixColumns = Math.max(
      0,
      ...items.map((item) => (Array.isArray(item) ? item.length : 1)),
    );
    const multiplier =
      binding.type === 'repeat-range' && binding.axis === 'both' && items.every(Array.isArray)
        ? matrixRows * matrixColumns
        : items.length;
    const height = binding.range.end.row - binding.range.start.row + 1;
    const width = binding.range.end.column - binding.range.start.column + 1;
    if (
      binding.type === 'repeat-page' ||
      (binding.type === 'repeat-range' && binding.axis === 'vertical')
    ) {
      estimatedRows += height * multiplier;
    } else if (
      binding.type === 'repeat-columns' ||
      (binding.type === 'repeat-range' && binding.axis === 'horizontal')
    ) {
      estimatedColumns += width * multiplier;
    } else if (binding.type === 'repeat-range' && binding.axis === 'both') {
      estimatedRows += height * matrixRows;
      estimatedColumns += width * matrixColumns;
    }
    return count + source * Math.max(0, multiplier - 1);
  }, sourceCells);
  const generatedSheets =
    nestedColumns.document.workbook.sheets.length +
    estimates
      .filter(({ binding }) => binding.type === 'repeat-sheet')
      .reduce((sum, { items }) => sum + items.length, 0);
  if (
    estimatedCells > limits.maxExpandedCells ||
    estimatedRows > limits.maxExpandedRows ||
    estimatedColumns > (limits.maxExpandedColumns ?? 16_384) ||
    generatedSheets > (limits.maxGeneratedSheets ?? 256)
  ) {
    return freeze({
      diagnostics: [
        error('EXPANSION_LIMIT_EXCEEDED', 'Advanced expansion exceeds configured limits'),
      ],
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  let document = nestedColumns.document;
  const mappings: StructuralMapping[] = [...nested.mappings, ...nestedColumns.mappings];
  const forcedPageBreaks = new Map(
    [...nested.breaks].map(([sheetId, values]) => [sheetId, [...values]]),
  );
  const generatedNames = new Set(document.workbook.sheets.map(({ name }) => name));
  const generatedIds = new Set(document.workbook.sheets.map(({ id }) => id));
  for (const { binding, items } of estimates.sort(
    (left, right) =>
      right.binding.range.start.row - left.binding.range.start.row ||
      right.binding.range.start.column - left.binding.range.start.column,
  )) {
    if (signal?.aborted) {
      return freeze({
        diagnostics: [error('RENDER_ABORTED', 'Template rendering was aborted')],
        structuralMappings: [],
        forcedPageBreaks: new Map(),
      });
    }
    if (binding.type === 'repeat-sheet') {
      const source = document.workbook.sheets.find(({ id }) => id === binding.range.sheetId);
      if (source === undefined) continue;
      const newSheets: Sheet[] = [];
      const sheetValues = compiled.ir.bindings.filter(
        (candidate): candidate is Extract<TemplateIRBinding, { readonly type: 'value' }> =>
          candidate.type === 'value' &&
          candidate.target.sheetId === source.id &&
          candidate.target.row >= binding.range.start.row &&
          candidate.target.row <= binding.range.end.row &&
          candidate.target.column >= binding.range.start.column &&
          candidate.target.column <= binding.range.end.column,
      );
      for (const [index, item] of items.entries()) {
        const nameValue =
          binding.name === undefined
            ? undefined
            : evaluateTemplateExpression(
                binding.name,
                {
                  root: data,
                  item,
                  index,
                  first: index === 0,
                  last: index === items.length - 1,
                },
                formatters,
              );
        const name = safeSheetName(nameValue, `${source.name} ${index + 1}`);
        if (generatedNames.has(name)) {
          return freeze({
            diagnostics: [
              error(
                'DUPLICATE_GENERATED_SHEET_NAME',
                `Generated sheet name ${name} is duplicated`,
                binding.id,
              ),
            ],
            structuralMappings: [],
            forcedPageBreaks: new Map(),
          });
        }
        generatedNames.add(name);
        let id = `${String(source.id)}~${String(binding.id)}~${index + 1}`;
        while (generatedIds.has(id as never)) id += '_';
        generatedIds.add(id as never);
        let cells = [...source.cells];
        const scope = {
          root: data,
          item,
          index,
          first: index === 0,
          last: index === items.length - 1,
        };
        for (const valueBinding of sheetValues) {
          const existing = cells.find(
            ({ row, column }) =>
              row === valueBinding.target.row && column === valueBinding.target.column,
          );
          const value = evaluateTemplateExpression(valueBinding.expression, scope, formatters);
          cells = [
            ...cells.filter(
              ({ row, column }) =>
                row !== valueBinding.target.row || column !== valueBinding.target.column,
            ),
            {
              row: valueBinding.target.row,
              column: valueBinding.target.column,
              cell: { ...existing?.cell, input: valueInput(value) },
            },
          ];
        }
        newSheets.push(
          freeze({
            ...source,
            id: id as never,
            name,
            cells: cells.sort((left, right) => left.row - right.row || left.column - right.column),
          }),
        );
        mappings.push({
          bindingId: binding.id,
          itemIndex: index,
          source: binding.range,
          generated: { ...binding.range, sheetId: id as never },
          ...(binding.objectPolicy === undefined ? {} : { objectPolicy: binding.objectPolicy }),
        });
      }
      document = freeze({
        ...document,
        workbook: {
          ...document.workbook,
          sheets: [...document.workbook.sheets, ...newSheets],
        },
      });
      continue;
    }
    if (binding.type === 'subtemplate') {
      const registered = compiled.ir.subtemplates?.find(({ id }) => id === binding.templateId);
      if (registered === undefined) {
        return freeze({
          diagnostics: [
            error(
              'SUBTEMPLATE_NOT_FOUND',
              `Subtemplate ${binding.templateId} is not registered`,
              binding.id,
            ),
          ],
          structuralMappings: [],
          forcedPageBreaks: new Map(),
        });
      }
      const sheetIndex = document.workbook.sheets.findIndex(
        ({ id }) => id === binding.range.sheetId,
      );
      const context = evaluateTemplateExpression(binding.source, { root: data }, formatters);
      const subtemplateMap = new Map(
        (compiled.ir.subtemplates ?? []).map((template) => [template.id, template]),
      );
      const childCompilation = compileSpreadsheetTemplate(document, registered, {
        subtemplates: subtemplateMap,
        limits,
      });
      if (childCompilation.template === undefined) {
        return freeze({
          diagnostics: childCompilation.diagnostics,
          structuralMappings: [],
          forcedPageBreaks: new Map(),
        });
      }
      const childExpansion = expandAdvancedTemplate(
        childCompilation.template,
        context,
        limits,
        formatters,
        signal,
      );
      if (childExpansion.document === undefined) {
        return freeze({
          diagnostics: childExpansion.diagnostics,
          structuralMappings: [],
          forcedPageBreaks: new Map(),
        });
      }
      if (sheetIndex >= 0) {
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
        if (sourceSheet === undefined) continue;
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
              column:
                entry.column > binding.range.end.column ? entry.column + columnDelta : entry.column,
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
        document = freeze({
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
        });
      }
      mappings.push({
        bindingId: binding.id,
        itemIndex: 0,
        source: binding.range,
        generated: binding.range,
        ...(binding.objectPolicy === undefined ? {} : { objectPolicy: binding.objectPolicy }),
      });
      continue;
    }
    const sheetIndex = document.workbook.sheets.findIndex(({ id }) => id === binding.range.sheetId);
    if (sheetIndex < 0) continue;
    const preserveEmpty = items.length === 0 && binding.empty === 'keep-template-row';
    const copyCount = preserveEmpty ? 1 : items.length;
    const rowCopies =
      binding.type === 'repeat-range' && binding.axis === 'both'
        ? copyCount
        : binding.type === 'repeat-columns' ||
            (binding.type === 'repeat-range' && binding.axis === 'horizontal')
          ? 1
          : copyCount;
    const columnCopies =
      binding.type === 'repeat-range' && binding.axis === 'both'
        ? preserveEmpty
          ? 1
          : Math.max(0, ...items.map((item) => (Array.isArray(item) ? item.length : 1)))
        : binding.type === 'repeat-columns' ||
            (binding.type === 'repeat-range' && binding.axis === 'horizontal')
          ? copyCount
          : 1;
    const cloned = cloneRange(
      document.workbook.sheets[sheetIndex]!,
      binding.range,
      rowCopies,
      columnCopies,
      items,
      binding.type === 'repeat-columns'
        ? 'horizontal'
        : binding.type === 'repeat-range'
          ? binding.axis
          : 'vertical',
      compiled.ir.bindings.filter(
        (candidate): candidate is Extract<TemplateIRBinding, { readonly type: 'value' }> =>
          candidate.type === 'value' &&
          candidate.target.sheetId === binding.range.sheetId &&
          candidate.target.row >= binding.range.start.row &&
          candidate.target.row <= binding.range.end.row &&
          candidate.target.column >= binding.range.start.column &&
          candidate.target.column <= binding.range.end.column,
      ),
      data,
      formatters,
    );
    const height = binding.range.end.row - binding.range.start.row + 1;
    const width = binding.range.end.column - binding.range.start.column + 1;
    const rowDelta = height * (rowCopies - 1);
    const columnDelta = width * (columnCopies - 1);
    for (let index = 0; index < mappings.length; index += 1) {
      const mapping = mappings[index]!;
      if (mapping.generated.sheetId !== binding.range.sheetId) continue;
      const shiftPoint = (point: { readonly row: number; readonly column: number }) => ({
        row: point.row > binding.range.end.row ? point.row + rowDelta : point.row,
        column: point.column > binding.range.end.column ? point.column + columnDelta : point.column,
      });
      mappings[index] = {
        ...mapping,
        generated: {
          ...mapping.generated,
          start: shiftPoint(mapping.generated.start),
          end: shiftPoint(mapping.generated.end),
        },
      };
    }
    document = freeze({
      ...document,
      workbook: {
        ...document.workbook,
        sheets: document.workbook.sheets.map((sheet, index) =>
          index === sheetIndex ? cloned.sheet : sheet,
        ),
      },
    });
    cloned.mappings.forEach((generated, index) => {
      mappings.push({
        bindingId: binding.id,
        itemIndex: index,
        source: binding.range,
        generated,
        ...(binding.objectPolicy === undefined ? {} : { objectPolicy: binding.objectPolicy }),
      });
    });
    if (binding.type === 'repeat-page') {
      const breaks = forcedPageBreaks.get(binding.range.sheetId) ?? [];
      for (let index = 1; index < items.length; index += 1) {
        breaks.push(binding.range.start.row + index * height);
      }
      forcedPageBreaks.set(binding.range.sheetId, breaks);
    }
  }
  const pageCount = [...forcedPageBreaks.values()].reduce(
    (count, breaks) => count + breaks.length + 1,
    0,
  );
  if (pageCount > limits.maxPages) {
    return freeze({
      diagnostics: [error('EXPANSION_LIMIT_EXCEEDED', 'Advanced expansion exceeds page limits')],
      structuralMappings: [],
      forcedPageBreaks: new Map(),
    });
  }
  return freeze({
    document,
    diagnostics: base.diagnostics,
    structuralMappings: mappings,
    forcedPageBreaks,
  });
}
