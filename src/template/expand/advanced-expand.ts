import type {
  Cell,
  Diagnostic,
  DocumentCellRange,
  Sheet,
  SpreadsheetDocument,
} from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';
import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type { CompiledTemplate, RenderLimits, TemplateIRBinding } from '../model';
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
  const base = expandTemplate(
    compiled.sourceDocument,
    compiled.ir.template,
    tp1Bindings,
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
    base.document.workbook.sheets.length +
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
  let document = base.document;
  const mappings: StructuralMapping[] = [];
  const forcedPageBreaks = new Map<string, number[]>();
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
