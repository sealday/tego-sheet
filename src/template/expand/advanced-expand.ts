import type {
  Cell,
  Diagnostic,
  DocumentCellRange,
  Sheet,
  SpreadsheetDocument,
} from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';
import {
  compileTemplateExpression,
  evaluateTemplateExpression,
  type TemplateFormatterRegistry,
} from '../expression';
import type {
  CompiledTemplate,
  RenderLimits,
  TemplateIRBinding,
  TemplateRegionNode,
} from '../model';
import { expandTemplate } from './expand';

export interface StructuralMapping {
  readonly bindingId: string;
  readonly itemIndex: number;
  readonly source: DocumentCellRange;
  readonly generated: DocumentCellRange;
  readonly objectPolicy?: 'per-item' | 'shared' | 'forbidden';
}

export interface AdvancedExpansionResult {
  readonly document?: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly structuralMappings: readonly StructuralMapping[];
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
      for (const entry of sourceCells) {
        copies.push({
          ...entry,
          row: entry.row + rowDelta,
          column: entry.column + columnDelta,
          cell: translatedCell(entry.cell, rowDelta, columnDelta),
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

function bindingByType(
  compiled: CompiledTemplate,
  bindingId: string,
): TemplateIRBinding['type'] | undefined {
  return compiled.ir.bindings.find(({ id }) => id === bindingId)?.type;
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
  const isVerticalTree = (node: TemplateRegionNode): boolean =>
    byId.get(node.bindingId)?.type === 'repeat-rows' && node.children.every(isVerticalTree);
  const roots = (compiled.ir.regionTree ?? []).filter(isVerticalTree);
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
      readonly height: number;
    }

    const renderNode = (
      node: TemplateRegionNode,
      parentScope: Scope,
      destinationStart: number,
    ): Fragment => {
      const binding = byId.get(node.bindingId) as
        | Extract<TemplateIRBinding, { readonly type: 'repeat-rows' }>
        | undefined;
      if (binding === undefined) return { cells: [], rows: [], height: 0 };
      const resolved = evaluateTemplateExpression(binding.source, parentScope, formatters);
      const items = Array.isArray(resolved) ? resolved : [];
      const copies =
        items.length === 0 && binding.empty === 'keep-template-row' ? [undefined] : items;
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      let height = 0;
      for (const [index, item] of copies.entries()) {
        if (signal?.aborted) return { cells: [], rows: [], height: 0 };
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
        height += fragment.height;
      }
      return { cells, rows, height };
    };

    const renderRange = (
      range: DocumentCellRange,
      children: readonly TemplateRegionNode[],
      scope: Scope,
      destinationStart: number,
    ): Fragment => {
      const cells: Sheet['cells'][number][] = [];
      const rows: Sheet['rows'][number][] = [];
      const sortedChildren = [...children].sort(
        (left, right) => left.range.start.row - right.range.start.row,
      );
      let childIndex = 0;
      let sourceRow = range.start.row;
      let height = 0;
      while (sourceRow <= range.end.row) {
        const child = sortedChildren[childIndex];
        if (child !== undefined && child.range.start.row === sourceRow) {
          const fragment = renderNode(child, scope, destinationStart + height);
          cells.push(...fragment.cells);
          rows.push(...fragment.rows);
          height += fragment.height;
          sourceRow = child.range.end.row + 1;
          childIndex += 1;
          continue;
        }
        const generatedRow = destinationStart + height;
        const rowDelta = generatedRow - sourceRow;
        const sourceCells = sourceSheet.cells.filter(({ row }) => row === sourceRow);
        const targets = valueBindings.filter(
          ({ target }) => target.sheetId === sourceSheet.id && target.row === sourceRow,
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
      return { cells, rows, height };
    };

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
      merges: sourceSheet.merges
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
      binding.type === 'repeat-columns' ||
      binding.type === 'repeat-range' ||
      binding.type === 'repeat-page' ||
      binding.type === 'repeat-sheet' ||
      binding.type === 'subtemplate',
  );
  const tp1Bindings = compiled.ir.bindings.filter(
    (binding) =>
      binding.type === 'value' ||
      binding.type === 'repeat-rows' ||
      binding.type === 'conditional-range',
  );
  const nestedRanges = (compiled.ir.regionTree ?? [])
    .filter((node) => bindingByType(compiled, node.bindingId) === 'repeat-rows')
    .map(({ range }) => range);
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
  const estimates = advanced.map((binding) => ({
    binding,
    items: collection(binding, { root: data }, formatters),
  }));
  const sourceCells = compiled.sourceDocument.workbook.sheets.reduce(
    (count, sheet) => count + sheet.cells.length,
    0,
  );
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
    const multiplier =
      binding.type === 'repeat-range' && binding.axis === 'both' && items.every(Array.isArray)
        ? items.reduce((sum, row) => sum + (row as readonly unknown[]).length, 0)
        : items.length;
    return count + source * Math.max(0, multiplier - 1);
  }, sourceCells);
  const generatedSheets =
    nested.document.workbook.sheets.length +
    estimates
      .filter(({ binding }) => binding.type === 'repeat-sheet')
      .reduce((sum, { items }) => sum + items.length, 0);
  if (
    estimatedCells > limits.maxExpandedCells ||
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
  let document = nested.document;
  const mappings: StructuralMapping[] = [...nested.mappings];
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
        newSheets.push(freeze({ ...source, id: id as never, name }));
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
      if (sheetIndex >= 0) {
        const sheet = document.workbook.sheets[sheetIndex]!;
        const childValues = registered.bindings.filter(
          (
            candidate,
          ): candidate is Extract<
            (typeof registered.bindings)[number],
            { readonly type: 'value' }
          > => candidate.type === 'value',
        );
        const minimumRow = Math.min(
          ...childValues.map(({ target }) => target.row),
          binding.range.start.row,
        );
        const minimumColumn = Math.min(
          ...childValues.map(({ target }) => target.column),
          binding.range.start.column,
        );
        let cells = [...sheet.cells];
        for (const child of childValues) {
          const row = binding.range.start.row + child.target.row - minimumRow;
          const column = binding.range.start.column + child.target.column - minimumColumn;
          const value = evaluateTemplateExpression(
            compileTemplateExpression(child.expression),
            { root: data, item: context },
            formatters,
          );
          const existing = cells.find((entry) => entry.row === row && entry.column === column);
          cells = [
            ...cells.filter((entry) => entry.row !== row || entry.column !== column),
            {
              row,
              column,
              cell: { ...existing?.cell, input: valueInput(value) },
            },
          ];
        }
        document = freeze({
          ...document,
          workbook: {
            ...document.workbook,
            sheets: document.workbook.sheets.map((candidate, index) =>
              index === sheetIndex
                ? {
                    ...candidate,
                    cells: cells.sort(
                      (left, right) => left.row - right.row || left.column - right.column,
                    ),
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
    const rowCopies =
      binding.type === 'repeat-range' && binding.axis === 'both'
        ? Math.max(1, items.length)
        : binding.type === 'repeat-columns' ||
            (binding.type === 'repeat-range' && binding.axis === 'horizontal')
          ? 1
          : Math.max(1, items.length);
    const columnCopies =
      binding.type === 'repeat-range' && binding.axis === 'both'
        ? Math.max(1, ...items.map((item) => (Array.isArray(item) ? item.length : 1)))
        : binding.type === 'repeat-columns' ||
            (binding.type === 'repeat-range' && binding.axis === 'horizontal')
          ? Math.max(1, items.length)
          : 1;
    const cloned = cloneRange(
      document.workbook.sheets[sheetIndex]!,
      binding.range,
      rowCopies,
      columnCopies,
    );
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
      const height = binding.range.end.row - binding.range.start.row + 1;
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
