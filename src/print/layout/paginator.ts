import type { Diagnostic } from '../../document';
import type { PrintScale } from '../../template/model';

export interface PaginationTarget {
  readonly id: string;
  readonly rows: readonly number[];
  readonly width: number;
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
  const widthScale =
    target.width <= 0
      ? 1
      : (availableWidth * ('pages' in input.scale ? input.scale.pages : 1)) / target.width;
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
          scale,
        }),
      );
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
      if (rowHeight > availableHeight) {
        emit(row - 1);
        diagnostics.push(
          diagnostic('ROW_EXCEEDS_PAGE', `Row ${row} in ${target.id} exceeds the printable page`),
        );
        start = row + 1;
        height = 0;
        continue;
      }
      if ((breaks.has(row) && row > start) || height + rowHeight > availableHeight) {
        emit(row - 1);
        start = row;
        height = 0;
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
