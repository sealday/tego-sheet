import type { Diagnostic } from '../../document';
import type { PrintScale } from '../../template/model';

export interface PaginationTarget {
  readonly id: string;
  readonly rows: readonly number[];
  readonly columns: readonly number[];
  readonly repeatRows?: readonly number[];
  readonly repeatColumns?: readonly number[];
}

export interface PaginationBreak {
  readonly targetId: string;
  readonly beforeRow: number;
}

export interface PaginationInput {
  readonly targets: readonly PaginationTarget[];
  readonly paper: { readonly width: number; readonly height: number };
  readonly margins: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly scale: PrintScale;
  readonly manualBreaks: readonly PaginationBreak[];
  readonly maxPages: number;
  readonly signal?: AbortSignal;
  readonly deadline?: number;
}

export interface PaginationPage {
  readonly id: string;
  readonly index: number;
  readonly targetId: string;
  readonly width: number;
  readonly height: number;
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly columnStart: number;
  readonly columnEnd: number;
  readonly scale: number;
}

export interface PaginationResult {
  readonly pages: readonly PaginationPage[];
  readonly diagnostics: readonly Diagnostic[];
}

function diagnostic(code: string, message: string): Diagnostic {
  return Object.freeze({
    code,
    severity: 'error',
    domain: 'layout',
    stage: 'layout',
    message,
  });
}

function aborted(input: PaginationInput): boolean {
  return input.signal?.aborted === true;
}

function scaleFor(input: PaginationInput, target: PaginationTarget): number {
  const availableWidth = input.paper.width - input.margins.left - input.margins.right;
  const availableHeight = input.paper.height - input.margins.top - input.margins.bottom;
  if (input.scale.type === 'fixed') return input.scale.value;
  const targetWidth = target.columns.reduce((sum, width) => sum + width, 0);
  const widthScale =
    targetWidth <= 0
      ? 1
      : (availableWidth * ('pages' in input.scale ? input.scale.pages : 1)) / targetWidth;
  if (input.scale.type === 'fit-width') return Math.min(1, widthScale);
  const totalHeight = target.rows.reduce((sum, height) => sum + height, 0);
  const heightScale = totalHeight <= 0 ? 1 : availableHeight / totalHeight;
  return Math.min(1, widthScale, heightScale);
}

/** Paginates explicit targets without consulting viewport or mutable renderer state. */
export function paginateTemplateTargets(input: PaginationInput): PaginationResult {
  const availableHeight = input.paper.height - input.margins.top - input.margins.bottom;
  if (availableHeight <= 0 || input.paper.width - input.margins.left - input.margins.right <= 0) {
    return Object.freeze({
      pages: Object.freeze([]),
      diagnostics: Object.freeze([diagnostic('INVALID_PRINT_TARGET', 'Margins consume the page')]),
    });
  }
  const diagnostics: Diagnostic[] = [];
  const pages: PaginationPage[] = [];
  let terminalDiagnostic: Diagnostic | undefined;
  for (const target of input.targets) {
    if (aborted(input)) {
      terminalDiagnostic = diagnostic('RENDER_ABORTED', 'Template rendering was aborted');
      break;
    }
    if (input.deadline !== undefined && Date.now() > input.deadline) {
      terminalDiagnostic = diagnostic('LAYOUT_TIME_EXCEEDED', 'Template layout exceeded its limit');
      break;
    }
    const scale = scaleFor(input, target);
    const availableTargetHeight =
      availableHeight - (target.repeatRows ?? []).reduce((sum, height) => sum + height * scale, 0);
    const availableTargetWidth =
      input.paper.width -
      input.margins.left -
      input.margins.right -
      (target.repeatColumns ?? []).reduce((sum, width) => sum + width * scale, 0);
    const columnSegments: { readonly start: number; readonly end: number }[] = [];
    let columnStart = 0;
    let columnWidth = 0;
    for (let column = 0; column < target.columns.length; column += 1) {
      const width = target.columns[column]! * scale;
      const pageWidth =
        columnStart === 0
          ? input.paper.width - input.margins.left - input.margins.right
          : availableTargetWidth;
      if (pageWidth <= 0 || width > pageWidth) {
        diagnostics.push(
          diagnostic(
            'COLUMN_EXCEEDS_PAGE',
            `Column ${column} in ${target.id} exceeds the printable page`,
          ),
        );
        continue;
      }
      if (columnWidth + width > pageWidth && column > columnStart) {
        columnSegments.push({ start: columnStart, end: column - 1 });
        columnStart = column;
        columnWidth = 0;
        if (availableTargetWidth <= 0 || width > availableTargetWidth) {
          diagnostics.push(
            diagnostic(
              'COLUMN_EXCEEDS_PAGE',
              `Column ${column} in ${target.id} exceeds the printable page`,
            ),
          );
          columnStart = column + 1;
          continue;
        }
      }
      columnWidth += width;
    }
    if (target.columns.length > 0 && columnStart < target.columns.length) {
      columnSegments.push({ start: columnStart, end: target.columns.length - 1 });
    }
    const breaks = new Set(
      input.manualBreaks
        .filter(({ targetId }) => targetId === target.id)
        .map(({ beforeRow }) => beforeRow),
    );
    let start = 0;
    let height = 0;
    let pageWithinTarget = 0;
    const emit = (end: number): void => {
      if (end < start || terminalDiagnostic !== undefined) return;
      for (const columns of columnSegments) {
        if (pages.length >= input.maxPages) {
          terminalDiagnostic = diagnostic(
            'PAGE_LIMIT_EXCEEDED',
            `Pagination exceeds ${input.maxPages} pages`,
          );
          return;
        }
        pages.push(
          Object.freeze({
            id: `${target.id}:${pageWithinTarget++}`,
            index: pages.length,
            targetId: target.id,
            width: input.paper.width,
            height: input.paper.height,
            rowStart: start,
            rowEnd: end,
            columnStart: columns.start,
            columnEnd: columns.end,
            scale,
          }),
        );
      }
    };
    for (let row = 0; row < target.rows.length; row += 1) {
      if (aborted(input)) {
        terminalDiagnostic = diagnostic('RENDER_ABORTED', 'Template rendering was aborted');
        break;
      }
      if (input.deadline !== undefined && Date.now() > input.deadline) {
        terminalDiagnostic = diagnostic(
          'LAYOUT_TIME_EXCEEDED',
          'Template layout exceeded its limit',
        );
        break;
      }
      const rowHeight = target.rows[row]! * scale;
      const pageHeight = start === 0 ? availableHeight : availableTargetHeight;
      if (pageHeight <= 0 || rowHeight > pageHeight) {
        emit(row - 1);
        diagnostics.push(
          diagnostic('ROW_EXCEEDS_PAGE', `Row ${row} in ${target.id} exceeds the printable page`),
        );
        start = row + 1;
        height = 0;
        continue;
      }
      if ((breaks.has(row) && row > start) || height + rowHeight > pageHeight) {
        emit(row - 1);
        start = row;
        height = 0;
        if (availableTargetHeight <= 0 || rowHeight > availableTargetHeight) {
          diagnostics.push(
            diagnostic('ROW_EXCEEDS_PAGE', `Row ${row} in ${target.id} exceeds the printable page`),
          );
          start = row + 1;
          continue;
        }
      }
      height += rowHeight;
    }
    emit(target.rows.length - 1);
    if (terminalDiagnostic !== undefined) break;
  }
  if (terminalDiagnostic !== undefined) {
    return Object.freeze({
      pages: Object.freeze([]),
      diagnostics: Object.freeze([terminalDiagnostic]),
    });
  }
  return Object.freeze({
    pages: Object.freeze(pages),
    diagnostics: Object.freeze(diagnostics),
  });
}
