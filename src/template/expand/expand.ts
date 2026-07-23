import type {
  Cell,
  Diagnostic,
  DocumentCellRange,
  Sheet,
  SpreadsheetDocument,
} from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';
import {
  evaluateTemplateExpression,
  TemplateExpressionError,
  type TemplateFormatterRegistry,
} from '../expression';
import type { RenderLimits, SpreadsheetTemplate, TemplateIRBinding } from '../model';

export interface ExpansionResult {
  readonly document?: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly insertedRows: ReadonlyMap<string, readonly RowInsertion[]>;
  readonly repeatPageBreaks: ReadonlyMap<string, readonly number[]>;
}

export interface RowInsertion {
  readonly afterSourceRow: number;
  readonly delta: number;
}

function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (value !== null && typeof value === 'object') {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

function cellInput(value: unknown): Cell['input'] {
  if (value === undefined || value === null) return { type: 'blank' };
  if (typeof value === 'string') return { type: 'string', value };
  if (typeof value === 'number' && Number.isFinite(value)) return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'string', value: JSON.stringify(value) };
}

function formatBindingValue(
  value: unknown,
  formatter: string | undefined,
  formatters: TemplateFormatterRegistry,
): unknown {
  if (formatter === undefined) return value;
  if (!Object.prototype.hasOwnProperty.call(formatters, formatter)) {
    throw new TemplateExpressionError('UNKNOWN_FORMATTER', `Unknown formatter: ${formatter}`);
  }
  const descriptor = Object.getOwnPropertyDescriptor(formatters, formatter);
  const callable = descriptor !== undefined && 'value' in descriptor ? descriptor.value : undefined;
  if (typeof callable !== 'function') {
    throw new TemplateExpressionError('UNKNOWN_FORMATTER', `Unknown formatter: ${formatter}`);
  }
  try {
    return callable(value);
  } catch {
    throw new TemplateExpressionError('FORMATTER_FAILED', `Formatter ${formatter} failed`);
  }
}

function translatedCell(cell: Cell, rowDelta: number): Cell {
  if (cell.input.type !== 'formula' || rowDelta === 0) return cell;
  try {
    return {
      ...cell,
      input: {
        type: 'formula',
        source: renderFormula(
          translateFormula(parseFormula(cell.input.source), { rowDelta, columnDelta: 0 }),
        ),
      },
    };
  } catch {
    return cell;
  }
}

function mapRow(row: number, insertions: readonly RowInsertion[]): number {
  return (
    row +
    insertions.reduce(
      (delta, insertion) => delta + (row > insertion.afterSourceRow ? insertion.delta : 0),
      0,
    )
  );
}

function removeRange(sheet: Sheet, range: DocumentCellRange): Sheet {
  const count = range.end.row - range.start.row + 1;
  const cells = sheet.cells
    .filter(({ row }) => row < range.start.row || row > range.end.row)
    .map((entry) => (entry.row > range.end.row ? { ...entry, row: entry.row - count } : entry));
  const rows = sheet.rows
    .filter(({ index }) => index < range.start.row || index > range.end.row)
    .map((row) => (row.index > range.end.row ? { ...row, index: row.index - count } : row));
  const merges = sheet.merges
    .filter((merge) => merge.end.row < range.start.row || merge.start.row > range.end.row)
    .map((merge) =>
      merge.start.row > range.end.row
        ? {
            start: { ...merge.start, row: merge.start.row - count },
            end: { ...merge.end, row: merge.end.row - count },
          }
        : merge,
    );
  return { ...sheet, cells, rows, merges };
}

function expandRepeat(
  sheet: Sheet,
  binding: Extract<TemplateIRBinding, { readonly type: 'repeat-rows' }>,
  items: readonly unknown[],
  data: unknown,
  formatters: TemplateFormatterRegistry,
  valueBindings: readonly Extract<TemplateIRBinding, { readonly type: 'value' }>[],
  diagnostics: Diagnostic[],
  signal?: AbortSignal,
): Sheet {
  const height = binding.range.end.row - binding.range.start.row + 1;
  const preserveTemplate = items.length === 0 && binding.empty === 'keep-template-row';
  const copies = preserveTemplate ? [undefined] : items;
  if (copies.length === 0) return removeRange(sheet, binding.range);
  const delta = height * (copies.length - 1);
  const existingSourceCells = sheet.cells.filter(
    ({ row }) => row >= binding.range.start.row && row <= binding.range.end.row,
  );
  const sourceCells = [...existingSourceCells];
  for (const valueBinding of preserveTemplate ? [] : valueBindings) {
    if (
      !sourceCells.some(
        ({ row, column }) =>
          row === valueBinding.target.row && column === valueBinding.target.column,
      )
    ) {
      sourceCells.push({
        row: valueBinding.target.row,
        column: valueBinding.target.column,
        cell: { input: { type: 'blank' } },
      });
    }
  }
  const before = sheet.cells.filter(({ row }) => row < binding.range.start.row);
  const after = sheet.cells
    .filter(({ row }) => row > binding.range.end.row)
    .map((entry) => ({ ...entry, row: entry.row + delta }));
  const repeated: Sheet['cells'][number][] = [];
  for (let itemIndex = 0; itemIndex < copies.length; itemIndex += 1) {
    if (signal?.aborted === true) break;
    const item = copies[itemIndex];
    const rowDelta = itemIndex * height;
    for (const entry of sourceCells) {
      const row = entry.row + rowDelta;
      const valueBinding = valueBindings.find(
        ({ target }) => target.row === entry.row && target.column === entry.column,
      );
      let cell = translatedCell(entry.cell, rowDelta);
      if (valueBinding !== undefined && !preserveTemplate) {
        let value = evaluateTemplateExpression(
          valueBinding.expression,
          {
            root: data,
            item,
            index: itemIndex,
            first: itemIndex === 0,
            last: itemIndex === copies.length - 1,
          },
          formatters,
        );
        if (value === undefined) {
          diagnostics.push({
            code: 'MISSING_DATA',
            severity: 'error',
            domain: 'template',
            stage: 'resolve',
            message: `Binding ${valueBinding.id} resolved to a missing value`,
            location: { bindingId: valueBinding.id },
          });
          value = null;
        }
        cell = {
          ...entry.cell,
          input: cellInput(formatBindingValue(value, valueBinding.formatter, formatters)),
        };
      }
      repeated.push({ ...entry, row, cell });
    }
  }
  const sourceRows = sheet.rows.filter(
    ({ index }) => index >= binding.range.start.row && index <= binding.range.end.row,
  );
  const rows = [
    ...sheet.rows.filter(({ index }) => index < binding.range.start.row),
    ...copies.flatMap((_, itemIndex) =>
      sourceRows.map((row) => ({ ...row, index: row.index + itemIndex * height })),
    ),
    ...sheet.rows
      .filter(({ index }) => index > binding.range.end.row)
      .map((row) => ({ ...row, index: row.index + delta })),
  ];
  const sourceMerges = sheet.merges.filter(
    ({ start, end }) => start.row >= binding.range.start.row && end.row <= binding.range.end.row,
  );
  const merges = [
    ...sheet.merges.filter(({ end }) => end.row < binding.range.start.row),
    ...copies.flatMap((_, itemIndex) =>
      sourceMerges.map((merge) => ({
        start: { ...merge.start, row: merge.start.row + itemIndex * height },
        end: { ...merge.end, row: merge.end.row + itemIndex * height },
      })),
    ),
    ...sheet.merges
      .filter(({ start }) => start.row > binding.range.end.row)
      .map((merge) => ({
        start: { ...merge.start, row: merge.start.row + delta },
        end: { ...merge.end, row: merge.end.row + delta },
      })),
  ];
  return { ...sheet, cells: [...before, ...repeated, ...after], rows, merges };
}

/** Expands scalar, repeat-row, and conditional bindings on an isolated document graph. */
export function expandTemplate(
  source: SpreadsheetDocument,
  template: SpreadsheetTemplate,
  bindings: readonly TemplateIRBinding[],
  data: unknown,
  formatters: TemplateFormatterRegistry,
  limits: RenderLimits,
  signal?: AbortSignal,
): ExpansionResult {
  const diagnostics: Diagnostic[] = [];
  const insertions = new Map<string, RowInsertion[]>();
  const repeatPageBreaks = new Map<string, number[]>();
  let expandedRows = 0;
  const sheets = source.workbook.sheets.map((sheet) => ({ ...sheet }));
  const repeats = bindings
    .filter(
      (binding): binding is Extract<TemplateIRBinding, { readonly type: 'repeat-rows' }> =>
        binding.type === 'repeat-rows',
    )
    .sort((left, right) => right.range.start.row - left.range.start.row);
  for (const binding of repeats) {
    if (signal?.aborted === true) {
      return freeze({
        diagnostics: [
          {
            code: 'RENDER_ABORTED',
            severity: 'error',
            domain: 'template',
            stage: 'expand',
            message: 'Template rendering was aborted',
          },
        ],
        insertedRows: insertions,
        repeatPageBreaks,
      });
    }
    const value = evaluateTemplateExpression(binding.source, { root: data }, formatters);
    const items = Array.isArray(value) ? value : [];
    const height = binding.range.end.row - binding.range.start.row + 1;
    const copies = items.length === 0 && binding.empty === 'keep-template-row' ? 1 : items.length;
    const preserveTemplate = items.length === 0 && binding.empty === 'keep-template-row';
    expandedRows += height * copies;
    if (expandedRows > limits.maxExpandedRows) {
      return freeze({
        diagnostics: [
          {
            code: 'EXPANSION_LIMIT_EXCEEDED',
            severity: 'error',
            domain: 'template',
            stage: 'expand',
            message: `Expanded rows exceed ${limits.maxExpandedRows}`,
            location: { bindingId: binding.id },
          },
        ],
        insertedRows: insertions,
        repeatPageBreaks,
      });
    }
    const sheetIndex = sheets.findIndex(({ id }) => id === binding.range.sheetId);
    if (sheetIndex < 0) continue;
    const valueBindings = bindings.filter(
      (candidate): candidate is Extract<TemplateIRBinding, { readonly type: 'value' }> =>
        candidate.type === 'value' &&
        candidate.target.sheetId === binding.range.sheetId &&
        candidate.target.row >= binding.range.start.row &&
        candidate.target.row <= binding.range.end.row,
    );
    const sourceCellCount =
      sheets[sheetIndex]!.cells.filter(
        ({ row }) => row >= binding.range.start.row && row <= binding.range.end.row,
      ).length +
      (preserveTemplate ? [] : valueBindings).filter(
        ({ target }) =>
          !sheets[sheetIndex]!.cells.some(
            ({ row, column }) => row === target.row && column === target.column,
          ),
      ).length;
    const existingCellCount = sheets.reduce(
      (count, candidate) => count + candidate.cells.length,
      0,
    );
    const replacedCellCount = sheets[sheetIndex]!.cells.filter(
      ({ row }) => row >= binding.range.start.row && row <= binding.range.end.row,
    ).length;
    const projectedCellCount = existingCellCount - replacedCellCount + sourceCellCount * copies;
    if (projectedCellCount > limits.maxExpandedCells) {
      return freeze({
        diagnostics: [
          {
            code: 'EXPANSION_LIMIT_EXCEEDED',
            severity: 'error',
            domain: 'template',
            stage: 'expand',
            message: `Expanded cells exceed ${limits.maxExpandedCells}`,
            location: { bindingId: binding.id },
          },
        ],
        insertedRows: insertions,
        repeatPageBreaks,
      });
    }
    sheets[sheetIndex] = expandRepeat(
      sheets[sheetIndex]!,
      binding,
      items,
      data,
      formatters,
      preserveTemplate ? [] : valueBindings,
      diagnostics,
      signal,
    );
    const delta = height * (copies - 1);
    const existingBreaks = repeatPageBreaks.get(binding.range.sheetId) ?? [];
    const shiftedBreaks = existingBreaks.map((row) =>
      row > binding.range.end.row ? row + delta : row,
    );
    if (binding.pageBreak === 'before-each-item') {
      for (let itemIndex = 1; itemIndex < copies; itemIndex += 1) {
        shiftedBreaks.push(binding.range.start.row + itemIndex * height);
      }
    }
    repeatPageBreaks.set(binding.range.sheetId, shiftedBreaks);
    const list = insertions.get(binding.range.sheetId) ?? [];
    list.push({ afterSourceRow: binding.range.end.row, delta });
    insertions.set(binding.range.sheetId, list);
  }
  for (const binding of bindings) {
    if (signal?.aborted === true) {
      return freeze({
        diagnostics: [
          {
            code: 'RENDER_ABORTED',
            severity: 'error',
            domain: 'template',
            stage: 'expand',
            message: 'Template rendering was aborted',
          },
        ],
        insertedRows: insertions,
        repeatPageBreaks,
      });
    }
    if (binding.type !== 'conditional-range') continue;
    if (evaluateTemplateExpression(binding.when, { root: data }, formatters)) continue;
    const sheetIndex = sheets.findIndex(({ id }) => id === binding.range.sheetId);
    if (sheetIndex < 0) continue;
    const sheetInsertions = insertions.get(binding.range.sheetId) ?? [];
    const mappedRange = {
      ...binding.range,
      start: { ...binding.range.start, row: mapRow(binding.range.start.row, sheetInsertions) },
      end: { ...binding.range.end, row: mapRow(binding.range.end.row, sheetInsertions) },
    };
    sheets[sheetIndex] = removeRange(sheets[sheetIndex]!, mappedRange);
    const list = insertions.get(binding.range.sheetId) ?? [];
    list.push({
      afterSourceRow: binding.range.end.row,
      delta: -(binding.range.end.row - binding.range.start.row + 1),
    });
    insertions.set(binding.range.sheetId, list);
  }
  for (const binding of bindings) {
    if (binding.type !== 'value') continue;
    if (
      repeats.some(
        ({ range }) =>
          range.sheetId === binding.target.sheetId &&
          binding.target.row >= range.start.row &&
          binding.target.row <= range.end.row,
      )
    ) {
      continue;
    }
    const sheetIndex = sheets.findIndex(({ id }) => id === binding.target.sheetId);
    if (sheetIndex < 0) continue;
    const sheet = sheets[sheetIndex]!;
    const row = mapRow(binding.target.row, insertions.get(binding.target.sheetId) ?? []);
    let value = evaluateTemplateExpression(binding.expression, { root: data }, formatters);
    if (value === undefined) {
      diagnostics.push({
        code: 'MISSING_DATA',
        severity: 'error',
        domain: 'template',
        stage: 'resolve',
        message: `Binding ${binding.id} resolved to a missing value`,
        location: { bindingId: binding.id },
      });
      value = null;
    }
    value = formatBindingValue(value, binding.formatter, formatters);
    const existing = sheet.cells.find(
      (entry) => entry.row === row && entry.column === binding.target.column,
    );
    const replacement = {
      row,
      column: binding.target.column,
      cell: { ...existing?.cell, input: cellInput(value) },
    };
    if (
      existing === undefined &&
      sheets.reduce((count, candidate) => count + candidate.cells.length, 0) + 1 >
        limits.maxExpandedCells
    ) {
      return freeze({
        diagnostics: [
          {
            code: 'EXPANSION_LIMIT_EXCEEDED',
            severity: 'error',
            domain: 'template',
            stage: 'expand',
            message: `Expanded cells exceed ${limits.maxExpandedCells}`,
            location: { bindingId: binding.id },
          },
        ],
        insertedRows: insertions,
        repeatPageBreaks,
      });
    }
    sheets[sheetIndex] = {
      ...sheet,
      cells: [
        ...sheet.cells.filter(
          (entry) => entry.row !== row || entry.column !== binding.target.column,
        ),
        replacement,
      ].sort((left, right) => left.row - right.row || left.column - right.column),
    };
  }
  const cellCount = sheets.reduce((count, sheet) => count + sheet.cells.length, 0);
  if (cellCount > limits.maxExpandedCells) {
    return freeze({
      diagnostics: [
        {
          code: 'EXPANSION_LIMIT_EXCEEDED',
          severity: 'error',
          domain: 'template',
          stage: 'expand',
          message: `Expanded cells exceed ${limits.maxExpandedCells}`,
        },
      ],
      insertedRows: insertions,
      repeatPageBreaks,
    });
  }
  return freeze({
    document: {
      ...source,
      workbook: { ...source.workbook, sheets },
      templates: source.templates,
    },
    diagnostics,
    insertedRows: insertions,
    repeatPageBreaks,
  });
}
