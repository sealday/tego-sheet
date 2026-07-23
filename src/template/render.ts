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
import { createPrintDisplayList, type PrintDisplayPageInput } from '../print';
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

function paginationTargets(resolved: readonly ResolvedTarget[]): readonly PaginationTarget[] {
  return resolved.map(({ id, sheet, range }) => ({
    id,
    rows: Array.from({ length: range.end.row - range.start.row + 1 }, (_, index) =>
      rowHeight(sheet, range.start.row + index),
    ),
    width: Array.from({ length: range.end.column - range.start.column + 1 }, (_, index) =>
      columnWidth(sheet, range.start.column + index),
    ).reduce((sum, width) => sum + width, 0),
  }));
}

function cellRect(
  sheet: Sheet,
  range: DocumentCellRange,
  page: PaginationPage,
  row: number,
  column: number,
  marginLeft: number,
  marginTop: number,
): { readonly x: number; readonly y: number; readonly width: number; readonly height: number } {
  const firstRow = range.start.row + page.rowStart;
  const x =
    marginLeft +
    Array.from({ length: column - range.start.column }, (_, index) =>
      columnWidth(sheet, range.start.column + index),
    ).reduce((sum, width) => sum + width, 0) *
      page.scale;
  const y =
    marginTop +
    Array.from({ length: row - firstRow }, (_, index) => rowHeight(sheet, firstRow + index)).reduce(
      (sum, height) => sum + height,
      0,
    ) *
      page.scale;
  return {
    x,
    y,
    width: columnWidth(sheet, column) * page.scale,
    height: rowHeight(sheet, row) * page.scale,
  };
}

function displayPages(
  pages: readonly PaginationPage[],
  resolved: readonly ResolvedTarget[],
  profile: TemplatePrintProfile,
  resolvePresentation: (sheetId: DocumentSheetId, row: number, column: number) => CellPresentation,
): readonly PrintDisplayPageInput[] {
  return pages.map((page) => {
    const target = resolved.find(({ id }) => id === page.targetId)!;
    const startRow = target.range.start.row + page.rowStart;
    const endRow = target.range.start.row + page.rowEnd;
    const cells = [];
    for (let row = startRow; row <= endRow; row += 1) {
      for (let column = target.range.start.column; column <= target.range.end.column; column += 1) {
        cells.push({
          rect: cellRect(
            target.sheet,
            target.range,
            page,
            row,
            column,
            profile.page.margins.left,
            profile.page.margins.top,
          ),
          presentation: resolvePresentation(target.sheet.id, row, column),
        });
      }
    }
    return { width: page.width, height: page.height, cells };
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
  const pagination = paginateTemplateTargets({
    targets: paginationTargets(resolvedTargets),
    paper: pageSize,
    margins: profile.page.margins,
    scale: profile.page.scale,
    manualBreaks: profile.manualBreaks.flatMap((pageBreak) =>
      resolvedTargets
        .filter(({ sheet }) => sheet.id === pageBreak.sheetId)
        .map((target) => ({
          targetId: target.id,
          beforeRow: Math.max(0, pageBreak.beforeRow - target.range.start.row),
        })),
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
  );
  const displayList = createPrintDisplayList({
    pages: pageInputs,
    fontMetrics: environment.fontMetrics,
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
