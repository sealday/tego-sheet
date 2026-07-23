import type { Diagnostic, DocumentCellRange, Sheet, SpreadsheetDocument } from '../../document';
import { compileSpreadsheetTemplate } from '../compiler';
import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type {
  CompiledTemplate,
  RenderLimits,
  TemplateIRBinding,
  TemplateRegionNode,
} from '../model';
import { cloneRange } from './advanced-clone-range';
import {
  expansionError as error,
  freeze,
  safeSheetName,
  translatedCell,
  valueInput,
  type ExpansionScope as Scope,
} from './advanced-internals';
import {
  createExpansionPlan,
  estimateAllocation,
  materializeAdvancedBindings,
} from './advanced-plan';
import { materializeSubtemplate } from './advanced-subtemplate';
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
  const isMixedTree = (node: TemplateRegionNode): boolean => {
    const type = byId.get(node.bindingId)?.type;
    return (
      (type === 'repeat-rows' || type === 'repeat-columns' || type === 'conditional-range') &&
      node.children.every(isMixedTree)
    );
  };
  const hasVerticalOrConditionalDescendant = (node: TemplateRegionNode): boolean =>
    node.children.some(
      (child) =>
        byId.get(child.bindingId)?.type !== 'repeat-columns' ||
        hasVerticalOrConditionalDescendant(child),
    );
  const roots = (compiled.ir.regionTree ?? []).filter((node) => {
    if (!isMixedTree(node)) return false;
    const type = byId.get(node.bindingId)?.type;
    return (
      type === 'repeat-rows' ||
      (type === 'repeat-columns' && hasVerticalOrConditionalDescendant(node))
    );
  });
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
              : childType === 'conditional-range'
                ? evaluateTemplateExpression(
                    (
                      byId.get(child.bindingId) as Extract<
                        TemplateIRBinding,
                        { readonly type: 'conditional-range' }
                      >
                    ).when,
                    scope,
                    formatters,
                  )
                  ? renderRange(child.range, child.children, scope, destinationStart + height)
                  : { cells: [], rows: [], merges: [], height: 0 }
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
          const mappingStart = mappings.length;
          const childFragment =
            childType === 'repeat-columns'
              ? renderHorizontalNode(
                  child,
                  scope,
                  destinationRow + height,
                  destinationColumn + (child.range.start.column - range.start.column),
                )
              : childType === 'conditional-range'
                ? evaluateTemplateExpression(
                    (
                      byId.get(child.bindingId) as Extract<
                        TemplateIRBinding,
                        { readonly type: 'conditional-range' }
                      >
                    ).when,
                    scope,
                    formatters,
                  )
                  ? renderHorizontalRange(
                      child.range,
                      child.children,
                      scope,
                      destinationRow + height,
                      destinationColumn + (child.range.start.column - range.start.column),
                    )
                  : { cells: [], rows: [], merges: [], height: 0, width: 0 }
                : renderNode(child, scope, destinationRow + height);
          const columnDelta = destinationColumn - range.start.column;
          if (childType !== 'repeat-columns') {
            for (let index = mappingStart; index < mappings.length; index += 1) {
              const mapping = mappings[index]!;
              mappings[index] = {
                ...mapping,
                generated: {
                  ...mapping.generated,
                  start: {
                    ...mapping.generated.start,
                    column: mapping.generated.start.column + columnDelta,
                  },
                  end: {
                    ...mapping.generated.end,
                    column: mapping.generated.end.column + columnDelta,
                  },
                },
              };
            }
          }
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
      const binding = byId.get(node.bindingId);
      if (binding?.type === 'conditional-range') {
        if (!evaluateTemplateExpression(binding.when, parentScope, formatters)) {
          return { rows: 0, cells: 0, exceeded: false };
        }
        const sourceCells = sourceSheet.cells.filter(
          ({ row, column }) =>
            row >= node.range.start.row &&
            row <= node.range.end.row &&
            column >= node.range.start.column &&
            column <= node.range.end.column,
        ).length;
        return {
          rows: node.range.end.row - node.range.start.row + 1,
          cells: sourceCells,
          exceeded: totalCells + sourceCells > limits.maxExpandedCells,
        };
      }
      if (binding?.type !== 'repeat-rows' && binding?.type !== 'repeat-columns') {
        return { rows: 0, cells: 0, exceeded: false };
      }
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
    const rootType = byId.get(root.bindingId)?.type;
    const fragment =
      rootType === 'repeat-columns'
        ? renderHorizontalNode(root, { root: data }, root.range.start.row, root.range.start.column)
        : renderNode(root, { root: data }, root.range.start.row);
    if (signal?.aborted) {
      return {
        diagnostics: [error('RENDER_ABORTED', 'Template rendering was aborted')],
        mappings: [],
        breaks: new Map(),
      };
    }
    const sourceHeight = root.range.end.row - root.range.start.row + 1;
    const sourceWidth = root.range.end.column - root.range.start.column + 1;
    const rowDelta = fragment.height - sourceHeight;
    const columnDelta =
      rootType === 'repeat-columns' ? (fragment as HorizontalFragment).width - sourceWidth : 0;
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
        ...sourceSheet.cells.filter(({ row, column }) =>
          rootType === 'repeat-columns'
            ? column < root.range.start.column
            : row < root.range.start.row,
        ),
        ...fragment.cells,
        ...sourceSheet.cells
          .filter(({ row, column }) =>
            rootType === 'repeat-columns'
              ? column > root.range.end.column
              : row > root.range.end.row,
          )
          .map((entry) =>
            rootType === 'repeat-columns'
              ? { ...entry, column: entry.column + columnDelta }
              : { ...entry, row: entry.row + rowDelta },
          ),
      ].sort((left, right) => left.row - right.row || left.column - right.column),
      rows: [
        ...sourceSheet.rows.filter(({ index }) => index < root.range.start.row),
        ...fragment.rows,
        ...sourceSheet.rows
          .filter(({ index }) => index > root.range.end.row)
          .map((row) => ({ ...row, index: row.index + rowDelta })),
      ],
      columns:
        rootType === 'repeat-columns'
          ? [
              ...sourceSheet.columns.filter(({ index }) => index < root.range.start.column),
              ...Array.from(
                { length: Math.max(0, (fragment as HorizontalFragment).width) },
                (_, offset) => {
                  const sourceIndex = root.range.start.column + (offset % sourceWidth);
                  const column = sourceSheet.columns.find(({ index }) => index === sourceIndex);
                  return column === undefined
                    ? undefined
                    : { ...column, index: root.range.start.column + offset };
                },
              ).filter((column): column is Sheet['columns'][number] => column !== undefined),
              ...sourceSheet.columns
                .filter(({ index }) => index > root.range.end.column)
                .map((column) => ({ ...column, index: column.index + columnDelta })),
            ]
          : sourceSheet.columns,
      merges: [
        ...sourceSheet.merges
          .filter(({ start, end }) =>
            rootType === 'repeat-columns'
              ? end.column < root.range.start.column || start.column > root.range.end.column
              : end.row < root.range.start.row || start.row > root.range.end.row,
          )
          .map((merge) =>
            rootType === 'repeat-columns' && merge.start.column > root.range.end.column
              ? {
                  start: { ...merge.start, column: merge.start.column + columnDelta },
                  end: { ...merge.end, column: merge.end.column + columnDelta },
                }
              : rootType !== 'repeat-columns' && merge.start.row > root.range.end.row
                ? {
                    start: { ...merge.start, row: merge.start.row + rowDelta },
                    end: { ...merge.end, row: merge.end.row + rowDelta },
                  }
                : merge,
          ),
        ...fragment.merges,
      ],
      ...(sourceSheet.rowCount === undefined ? {} : { rowCount: sourceSheet.rowCount + rowDelta }),
      ...(sourceSheet.columnCount === undefined
        ? {}
        : { columnCount: sourceSheet.columnCount + columnDelta }),
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
  const plan = createExpansionPlan(compiled);
  const base = expandTemplate(
    compiled.sourceDocument,
    compiled.ir.template,
    plan.tp1Bindings.filter((binding) => !plan.handledByNested(binding)),
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
  const estimates = materializeAdvancedBindings(plan.advanced, data, formatters);
  const allocation = estimateAllocation(compiled, estimates);
  const generatedSheets =
    nestedColumns.document.workbook.sheets.length +
    estimates
      .filter(({ binding }) => binding.type === 'repeat-sheet')
      .reduce((sum, { items }) => sum + items.length, 0);
  if (
    allocation.cells > limits.maxExpandedCells ||
    allocation.rows > limits.maxExpandedRows ||
    allocation.columns > (limits.maxExpandedColumns ?? 16_384) ||
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
  for (const { binding, items } of [...estimates].sort(
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
      const materialized = materializeSubtemplate(document, binding, registered, childExpansion);
      if (materialized === undefined) continue;
      document = materialized.document;
      mappings.push({
        bindingId: binding.id,
        itemIndex: 0,
        source: binding.range,
        generated: materialized.generatedRange,
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
