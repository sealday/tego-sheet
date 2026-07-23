import type {
  Diagnostic,
  DocumentCellRange,
  DocumentSheetId,
  Sheet,
  SpreadsheetDocument,
} from '../document';
import { createFormulaEngine } from '../formula';
import {
  createPresentationCache,
  createPresentationResolver,
  type CellPresentation,
} from '../presentation';
import {
  createPrintDisplayList,
  type PrintDisplayCommand,
  type PrintDisplayPageInput,
} from '../print';
import {
  paginateTemplateTargets,
  type PaginationPage,
  type PaginationTarget,
} from '../print/layout';
import { expandTemplate, type RowInsertion } from './expand';
import { TemplateExpressionError } from './expression';
import type {
  GeneratedDocument,
  RenderEnvironment,
  RenderLimits,
  RenderRequest,
  RenderResult,
  TemplatePrintProfile,
} from './model';

const DEFAULT_LIMITS: RenderLimits = Object.freeze({
  maxExpandedCells: 250_000,
  maxExpandedRows: 100_000,
  maxPages: 10_000,
  maxLayoutTimeMs: 5_000,
});

const PAPER = Object.freeze({
  A4: Object.freeze({ width: 794, height: 1123 }),
  A5: Object.freeze({ width: 559, height: 794 }),
  Letter: Object.freeze({ width: 816, height: 1056 }),
});

interface ResolvedTarget {
  readonly id: string;
  readonly sheet: Sheet;
  readonly range: DocumentCellRange;
}

function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (value !== null && typeof value === 'object') {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

function renderDiagnostic(code: string, message: string): Diagnostic {
  return freeze({
    code,
    severity: 'error',
    domain: 'template',
    stage: code === 'TEMPLATE_SOURCE_STALE' ? 'validate' : 'render',
    message,
  });
}

function mappedRow(row: number, insertions: readonly RowInsertion[]): number {
  return (
    row +
    insertions.reduce(
      (delta, insertion) => delta + (row > insertion.afterSourceRow ? insertion.delta : 0),
      0,
    )
  );
}

function usedRange(sheet: Sheet): DocumentCellRange {
  const maximumRow = Math.max(
    0,
    (sheet.rowCount ?? 1) - 1,
    ...sheet.cells.map(({ row }) => row),
    ...sheet.rows.map(({ index }) => index),
  );
  const maximumColumn = Math.max(
    0,
    (sheet.columnCount ?? 1) - 1,
    ...sheet.cells.map(({ column }) => column),
    ...sheet.columns.map(({ index }) => index),
  );
  return {
    sheetId: sheet.id,
    start: { row: 0, column: 0 },
    end: { row: maximumRow, column: maximumColumn },
  };
}

function transformRange(
  range: DocumentCellRange,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
): DocumentCellRange {
  const sheetInsertions = insertions.get(range.sheetId) ?? [];
  return {
    ...range,
    start: { ...range.start, row: mappedRow(range.start.row, sheetInsertions) },
    end: {
      ...range.end,
      row:
        range.end.row +
        sheetInsertions.reduce(
          (delta, insertion) =>
            delta + (range.end.row >= insertion.afterSourceRow ? Math.max(0, insertion.delta) : 0),
          0,
        ),
    },
  };
}

function targets(
  document: SpreadsheetDocument,
  profile: TemplatePrintProfile,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
): readonly ResolvedTarget[] {
  const output: ResolvedTarget[] = [];
  const append = (range: DocumentCellRange): void => {
    const transformed = transformRange(range, insertions);
    const sheet = document.workbook.sheets.find(({ id }) => id === transformed.sheetId);
    if (sheet === undefined) return;
    output.push({
      id: `${profile.id}:${output.length}:${sheet.id}:${transformed.start.row}-${transformed.end.row}:${transformed.start.column}-${transformed.end.column}`,
      sheet,
      range: transformed,
    });
  };
  for (const target of profile.targets) {
    if (target.type === 'sheet') {
      const sheet = document.workbook.sheets.find(({ id }) => id === target.sheetId);
      if (sheet !== undefined) append(usedRange(sheet));
    } else if (target.type === 'range') {
      append(target.range);
    } else {
      target.ranges.forEach(append);
    }
  }
  return output;
}

function rowHeight(sheet: Sheet, row: number): number {
  return sheet.rows.find(({ index }) => index === row)?.height ?? 20;
}

function columnWidth(sheet: Sheet, column: number): number {
  return sheet.columns.find(({ index }) => index === column)?.width ?? 100;
}

function paper(profile: TemplatePrintProfile): { readonly width: number; readonly height: number } {
  const definition =
    profile.page.paper.type === 'custom'
      ? { width: profile.page.paper.width, height: profile.page.paper.height }
      : PAPER[profile.page.paper.type];
  return profile.page.orientation === 'portrait'
    ? definition
    : { width: definition.height, height: definition.width };
}

function paginationTargets(
  resolved: readonly ResolvedTarget[],
  profile: TemplatePrintProfile,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
): readonly PaginationTarget[] {
  return resolved.map(({ id, sheet, range }) => ({
    id,
    rows: Array.from({ length: range.end.row - range.start.row + 1 }, (_, index) =>
      rowHeight(sheet, range.start.row + index),
    ),
    columns: Array.from({ length: range.end.column - range.start.column + 1 }, (_, index) =>
      columnWidth(sheet, range.start.column + index),
    ),
    ...(profile.repeatRows?.sheetId === sheet.id
      ? {
          repeatRows: (() => {
            const repeated = transformRange(profile.repeatRows, insertions);
            return Array.from({ length: repeated.end.row - repeated.start.row + 1 }, (_, index) =>
              rowHeight(sheet, repeated.start.row + index),
            );
          })(),
        }
      : {}),
    ...(profile.repeatColumns?.sheetId === sheet.id
      ? {
          repeatColumns: Array.from(
            { length: profile.repeatColumns.end.column - profile.repeatColumns.start.column + 1 },
            (_, index) => columnWidth(sheet, profile.repeatColumns!.start.column + index),
          ),
        }
      : {}),
  }));
}

function columnLabel(column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function bandText(source: string, page: number, pages: number, data: unknown, date: Date): string {
  return source.replace(/\{\{\s*([^}]+?)\s*\}\}/gu, (_, token: string) => {
    if (token === 'page') return String(page);
    if (token === 'pages') return String(pages);
    if (token === 'date') return date.toISOString().slice(0, 10);
    let value = data;
    for (const part of token.split('.')) {
      if (
        ['__proto__', 'prototype', 'constructor'].includes(part) ||
        value === null ||
        typeof value !== 'object' ||
        !Object.prototype.hasOwnProperty.call(value, part)
      ) {
        return '';
      }
      value = (value as Readonly<Record<string, unknown>>)[part];
    }
    return value === undefined || value === null ? '' : String(value);
  });
}

function bandCommands(
  profile: TemplatePrintProfile,
  page: PaginationPage,
  totalPages: number,
  data: unknown,
  date: Date,
): readonly PrintDisplayCommand[] {
  const commands: PrintDisplayCommand[] = [];
  for (const [band, y] of [
    [profile.header, Math.max(10, profile.page.margins.top / 2)],
    [profile.footer, page.height - Math.max(10, profile.page.margins.bottom / 2)],
  ] as const) {
    if (band === undefined) continue;
    for (const [source, align, x] of [
      [band.left, 'left', profile.page.margins.left],
      [band.center, 'center', page.width / 2],
      [band.right, 'right', page.width - profile.page.margins.right],
    ] as const) {
      if (source === undefined) continue;
      commands.push({
        kind: 'text',
        text: bandText(source, page.index + 1, totalPages, data, date),
        x,
        y,
        maxWidth: Math.max(0, page.width / 3),
        fontFamily: 'Arial',
        fontSize: 10,
        color: '#0a0a0a',
        horizontalAlign: align,
      });
    }
  }
  return commands;
}

function displayPages(
  pages: readonly PaginationPage[],
  resolved: readonly ResolvedTarget[],
  profile: TemplatePrintProfile,
  resolvePresentation: (sheetId: DocumentSheetId, row: number, column: number) => CellPresentation,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  data: unknown,
  date: Date,
): readonly PrintDisplayPageInput[] {
  return pages.map((page) => {
    const target = resolved.find(({ id }) => id === page.targetId)!;
    const startRow = target.range.start.row + page.rowStart;
    const endRow = target.range.start.row + page.rowEnd;
    const startColumn = target.range.start.column + page.columnStart;
    const endColumn = target.range.start.column + page.columnEnd;
    const bodyRows = Array.from({ length: endRow - startRow + 1 }, (_, index) => startRow + index);
    const bodyColumns = Array.from(
      { length: endColumn - startColumn + 1 },
      (_, index) => startColumn + index,
    );
    const titleRows =
      profile.repeatRows?.sheetId === target.sheet.id
        ? (() => {
            const range = transformRange(profile.repeatRows, insertions);
            return Array.from(
              { length: range.end.row - range.start.row + 1 },
              (_, index) => range.start.row + index,
            ).filter((row) => !bodyRows.includes(row));
          })()
        : [];
    const titleColumns =
      profile.repeatColumns?.sheetId === target.sheet.id
        ? Array.from(
            {
              length: profile.repeatColumns.end.column - profile.repeatColumns.start.column + 1,
            },
            (_, index) => profile.repeatColumns!.start.column + index,
          ).filter((column) => !bodyColumns.includes(column))
        : [];
    const rows = [...titleRows, ...bodyRows];
    const columns = [...titleColumns, ...bodyColumns];
    const headingSize = profile.showHeadings ? 20 : 0;
    const cells = [];
    for (const [rowIndex, row] of rows.entries()) {
      for (const [columnIndex, column] of columns.entries()) {
        const merge = target.sheet.merges.find(
          ({ start, end }) =>
            row >= start.row && row <= end.row && column >= start.column && column <= end.column,
        );
        if (merge !== undefined && (row !== merge.start.row || column !== merge.start.column)) {
          continue;
        }
        const mergeRows =
          merge === undefined
            ? [row]
            : rows.filter(
                (candidate) => candidate >= merge.start.row && candidate <= merge.end.row,
              );
        const mergeColumns =
          merge === undefined
            ? [column]
            : columns.filter(
                (candidate) => candidate >= merge.start.column && candidate <= merge.end.column,
              );
        cells.push({
          rect: {
            x:
              profile.page.margins.left +
              headingSize +
              columns
                .slice(0, columnIndex)
                .reduce((sum, candidate) => sum + columnWidth(target.sheet, candidate), 0) *
                page.scale,
            y:
              profile.page.margins.top +
              headingSize +
              rows
                .slice(0, rowIndex)
                .reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
                page.scale,
            width:
              mergeColumns.reduce(
                (sum, candidate) => sum + columnWidth(target.sheet, candidate),
                0,
              ) * page.scale,
            height:
              mergeRows.reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
              page.scale,
          },
          presentation: resolvePresentation(target.sheet.id, row, column),
        });
      }
    }
    const decorations: PrintDisplayCommand[] = [
      ...bandCommands(profile, page, pages.length, data, date),
    ];
    if (profile.showHeadings) {
      for (const [index, column] of columns.entries()) {
        decorations.push({
          kind: 'text',
          text: columnLabel(column),
          x:
            profile.page.margins.left +
            headingSize +
            columns
              .slice(0, index)
              .reduce((sum, candidate) => sum + columnWidth(target.sheet, candidate), 0) *
              page.scale +
            (columnWidth(target.sheet, column) * page.scale) / 2,
          y: profile.page.margins.top + headingSize / 2,
          maxWidth: columnWidth(target.sheet, column) * page.scale,
          fontFamily: 'Arial',
          fontSize: 10,
          color: '#0a0a0a',
          horizontalAlign: 'center',
        });
      }
      for (const [index, row] of rows.entries()) {
        decorations.push({
          kind: 'text',
          text: String(row + 1),
          x: profile.page.margins.left + headingSize / 2,
          y:
            profile.page.margins.top +
            headingSize +
            rows
              .slice(0, index)
              .reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
              page.scale +
            (rowHeight(target.sheet, row) * page.scale) / 2,
          maxWidth: headingSize,
          fontFamily: 'Arial',
          fontSize: 10,
          color: '#0a0a0a',
          horizontalAlign: 'center',
        });
      }
    }
    return {
      width: page.width,
      height: page.height,
      cells,
      decorations,
      showGridlines: profile.showGridlines,
    };
  });
}

function abortResult(): RenderResult {
  const diagnostic = renderDiagnostic('RENDER_ABORTED', 'Template rendering was aborted');
  return freeze({ diagnostics: [diagnostic] });
}

function isAborted(signal: AbortSignal | undefined): boolean {
  return signal?.aborted === true;
}

/** Executes expand → recalculate → present → paginate against one immutable source. */
export async function renderSpreadsheetTemplate(
  request: RenderRequest,
  environment: RenderEnvironment,
): Promise<RenderResult> {
  if (isAborted(request.signal)) return abortResult();
  if (request.template.sourceDocumentHash !== request.currentDocumentHash) {
    return freeze({
      diagnostics: [
        renderDiagnostic(
          'TEMPLATE_SOURCE_STALE',
          'Compiled template does not match the current document',
        ),
      ],
    });
  }
  const limits = Object.freeze({ ...DEFAULT_LIMITS, ...request.limits });
  const start = Date.now();
  const profile = request.template.ir.profiles.find(({ id }) => id === request.profileId);
  if (profile === undefined) {
    return freeze({
      diagnostics: [
        renderDiagnostic('INVALID_PRINT_TARGET', `Unknown print profile: ${request.profileId}`),
      ],
    });
  }
  let expansion;
  try {
    expansion = expandTemplate(
      request.template.sourceDocument,
      request.template.ir.template,
      request.template.ir.bindings,
      request.data,
      environment.formatters ?? {},
      limits,
      request.signal,
    );
  } catch (cause) {
    if (!(cause instanceof TemplateExpressionError)) throw cause;
    return freeze({
      diagnostics: [
        renderDiagnostic(
          cause.code,
          cause.code === 'FORMATTER_FAILED'
            ? 'A template formatter failed during rendering'
            : cause.message,
        ),
      ],
    });
  }
  if (isAborted(request.signal)) return abortResult();
  if (expansion.document === undefined) return freeze({ diagnostics: expansion.diagnostics });
  const expansionDiagnostics = expansion.diagnostics.map((diagnostic) =>
    request.missingValue === 'warning-and-blank' && diagnostic.code === 'MISSING_DATA'
      ? { ...diagnostic, severity: 'warning' as const }
      : diagnostic,
  );
  if (expansionDiagnostics.some(({ severity }) => severity === 'error')) {
    return freeze({ diagnostics: expansionDiagnostics });
  }
  const formulaEngine = createFormulaEngine();
  const formulaProgram = formulaEngine.compile(expansion.document);
  const calculation = formulaEngine.recalculate(formulaProgram, [], {
    locale: environment.locale,
    timeZone: environment.timeZone,
    dateSystem: environment.dateSystem,
    clock: { now: () => environment.clock.getTime() },
    tick: 0,
    functionRegistryVersion: 'builtin-1',
  });
  const resolvedTargets = targets(expansion.document, profile, expansion.insertedRows);
  const pageSize = paper(profile);
  const headingSize = profile.showHeadings ? 20 : 0;
  const pagination = paginateTemplateTargets({
    targets: paginationTargets(resolvedTargets, profile, expansion.insertedRows),
    paper: pageSize,
    margins: {
      ...profile.page.margins,
      top: profile.page.margins.top + headingSize,
      left: profile.page.margins.left + headingSize,
    },
    scale: profile.page.scale,
    manualBreaks: profile.manualBreaks
      .flatMap((pageBreak) =>
        resolvedTargets
          .filter(({ sheet }) => sheet.id === pageBreak.sheetId)
          .map((target) => ({
            targetId: target.id,
            beforeRow: Math.max(0, pageBreak.beforeRow - target.range.start.row),
          })),
      )
      .concat(
        resolvedTargets.flatMap((target) =>
          (expansion.repeatPageBreaks.get(target.sheet.id) ?? [])
            .filter(
              (beforeRow) =>
                beforeRow >= target.range.start.row && beforeRow <= target.range.end.row,
            )
            .map((beforeRow) => ({
              targetId: target.id,
              beforeRow: beforeRow - target.range.start.row,
            })),
        ),
      ),
    maxPages: limits.maxPages,
    signal: request.signal,
    deadline: start + limits.maxLayoutTimeMs,
  });
  if (Date.now() - start > limits.maxLayoutTimeMs || isAborted(request.signal)) {
    return isAborted(request.signal)
      ? abortResult()
      : freeze({
          diagnostics: [
            renderDiagnostic('LAYOUT_TIME_EXCEEDED', 'Template layout exceeded its time limit'),
          ],
        });
  }
  if (pagination.diagnostics.some(({ severity }) => severity === 'error')) {
    return freeze({ diagnostics: [...expansionDiagnostics, ...pagination.diagnostics] });
  }
  const cache = createPresentationCache({
    maximumEntries: Math.max(1, limits.maxExpandedCells),
    maximumBytes: Math.max(1_024, limits.maxExpandedCells * 1_024),
  });
  const presentation = createPresentationResolver({
    document: expansion.document,
    formulaProgram,
    formulaValues: calculation.values,
    cache,
    revisions: {
      document: 1,
      calculation: 1,
      condition: 1,
      style: 1,
      environment: 1,
    },
    environment: {
      locale: environment.locale,
      timeZone: environment.timeZone,
      dateSystem: environment.dateSystem,
      target: 'print',
    },
  });
  const pageInputs = displayPages(
    pagination.pages,
    resolvedTargets,
    profile,
    (sheetId, row, column) => presentation.resolve({ sheetId, row, column }),
    expansion.insertedRows,
    request.data,
    environment.clock,
  );
  const displayList = createPrintDisplayList({
    pages: pageInputs,
    fontMetrics: environment.fontMetrics,
    signal: request.signal,
    deadline: start + limits.maxLayoutTimeMs,
  });
  const diagnostics = [
    ...request.template.diagnostics,
    ...expansionDiagnostics,
    ...pagination.diagnostics,
    ...displayList.diagnostics,
  ];
  if (diagnostics.some(({ severity }) => severity === 'error')) {
    return freeze({ diagnostics });
  }
  const document: GeneratedDocument = freeze({
    workbook: expansion.document.workbook,
    print: {
      pages: pagination.pages.map((page) => ({
        id: page.id,
        index: page.index,
        targetId: page.targetId,
        width: page.width,
        height: page.height,
        rowStart: page.rowStart,
        rowEnd: page.rowEnd,
        columnStart: page.columnStart,
        columnEnd: page.columnEnd,
      })),
      displayList,
    },
    resources: {},
    diagnostics,
    metadata: {
      templateId: request.template.templateId,
      profileId: profile.id,
      sourceDocumentHash: request.template.sourceDocumentHash,
      locale: environment.locale,
      timeZone: environment.timeZone,
      generatedAt: environment.clock.toISOString(),
    },
  });
  return freeze({ document, diagnostics });
}
