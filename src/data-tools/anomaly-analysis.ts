import type { CellInput, Diagnostic, DocumentCellAddress, DocumentCellRange } from '../document';
import type { DocumentControllerSnapshot } from '../document-controller';
import { DataTransformError } from './errors';
import type { DataToolPreviewContext } from './transform-planner';

/** Read-only anomaly categories supported by the bounded analyzer. */
export type DataAnomalyCheck = 'blank' | 'error' | 'type-outlier';

/** One bounded read-only anomaly request. */
export interface DataAnomalyAnalysisRequest {
  /** Inclusive worksheet range to inspect. */
  readonly range: DocumentCellRange;
  /** Explicit anomaly categories to report. */
  readonly checks: readonly DataAnomalyCheck[];
  /** Optional expected input type; otherwise the dominant non-blank type is used. */
  readonly expectedType?: Exclude<CellInput['type'], 'blank'>;
}

/** Runtime budgets and optional calculation context for anomaly analysis. */
export interface DataAnomalyAnalysisOptions {
  /** Maximum cells inspected in one analysis. */
  readonly maxCells: number;
  /** Maximum findings returned before truncation. */
  readonly maxFindings: number;
  /** Cancels analysis without publishing a partial result. */
  readonly signal?: AbortSignal;
  /** Optional read-only host projections, including calculated error cells. */
  readonly context?: DataToolPreviewContext;
}

/** Immutable read-only anomaly result; it is intentionally not a commit plan. */
export interface DataAnomalyAnalysis {
  /** Exact number of cells inspected. */
  readonly inspectedCellCount: number;
  /** Whether findings were capped by the configured limit. */
  readonly truncated: boolean;
  /** Immutable anomaly diagnostics in row-major order. */
  readonly findings: readonly Diagnostic[];
}

function key(row: number, column: number): string {
  return `${row}:${column}`;
}

function addressKey(address: DocumentCellAddress): string {
  return `${address.sheetId}:${address.row}:${address.column}`;
}

function finding(
  code: string,
  message: string,
  cell: DocumentCellAddress,
  details?: Readonly<Record<string, string>>,
): Diagnostic {
  return Object.freeze({
    code,
    severity: 'warning',
    domain: 'data',
    stage: 'plan',
    message,
    location: Object.freeze({ cell: Object.freeze({ ...cell }) }),
    ...(details === undefined ? {} : { details: Object.freeze({ ...details }) }),
  });
}

function assertLimit(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new TypeError(`${label} must be a non-negative safe integer`);
  }
}

function throwIfAnalysisAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new DataTransformError('TRANSFORM_ABORTED', 'Data anomaly analysis was aborted');
  }
}

/** Analyzes data quality without creating commands, plans, history, or mutations. */
export async function analyzeDataAnomalies(
  snapshot: DocumentControllerSnapshot,
  request: DataAnomalyAnalysisRequest,
  options: DataAnomalyAnalysisOptions,
): Promise<DataAnomalyAnalysis> {
  assertLimit(options.maxCells, 'maxCells');
  assertLimit(options.maxFindings, 'maxFindings');
  throwIfAnalysisAborted(options.signal);
  const rowCount = request.range.end.row - request.range.start.row + 1;
  const columnCount = request.range.end.column - request.range.start.column + 1;
  if (rowCount < 1 || columnCount < 1 || rowCount > Math.floor(options.maxCells / columnCount)) {
    throw new DataTransformError(
      'TRANSFORM_TOO_LARGE',
      'Data anomaly analysis exceeds the configured cell limit',
    );
  }
  const sheet = snapshot.document.workbook.sheets.find(({ id }) => id === request.range.sheetId);
  if (sheet === undefined) {
    throw new DataTransformError('TRANSFORM_TOO_LARGE', 'Data anomaly sheet does not exist');
  }
  const cells = new Map(
    sheet.cells.map((entry) => [key(entry.row, entry.column), entry.cell.input]),
  );
  const errorCells = new Set((options.context?.errorCells ?? []).map(addressKey));
  let expectedType = request.expectedType;
  if (expectedType === undefined && request.checks.includes('type-outlier')) {
    const counts = new Map<Exclude<CellInput['type'], 'blank'>, number>();
    for (let row = request.range.start.row; row <= request.range.end.row; row += 1) {
      for (
        let column = request.range.start.column;
        column <= request.range.end.column;
        column += 1
      ) {
        const input = cells.get(key(row, column));
        if (input === undefined || input.type === 'blank') continue;
        const type = input.type;
        counts.set(type, (counts.get(type) ?? 0) + 1);
      }
    }
    expectedType = [...counts].sort(
      ([leftType, leftCount], [rightType, rightCount]) =>
        rightCount - leftCount || leftType.localeCompare(rightType),
    )[0]?.[0];
  }

  const findings: Diagnostic[] = [];
  let truncated = false;
  const add = (value: Diagnostic): void => {
    if (findings.length < options.maxFindings) findings.push(value);
    else truncated = true;
  };
  for (let row = request.range.start.row; row <= request.range.end.row; row += 1) {
    for (let column = request.range.start.column; column <= request.range.end.column; column += 1) {
      throwIfAnalysisAborted(options.signal);
      const cell = { sheetId: request.range.sheetId, row, column };
      const input = cells.get(key(row, column)) ?? ({ type: 'blank' } as const);
      const calculatedError = errorCells.has(addressKey(cell));
      if (request.checks.includes('error') && calculatedError) {
        add(finding('DATA_ERROR_ANOMALY', 'Calculated cell contains an error', cell));
        continue;
      }
      if (request.checks.includes('blank') && input.type === 'blank') {
        add(finding('DATA_BLANK_ANOMALY', 'Cell is blank', cell));
        continue;
      }
      if (
        request.checks.includes('type-outlier') &&
        input.type !== 'blank' &&
        expectedType !== undefined &&
        input.type !== expectedType
      ) {
        add(
          finding('DATA_TYPE_OUTLIER', 'Cell input type differs from the expected type', cell, {
            actualType: input.type,
            expectedType,
          }),
        );
      }
    }
  }
  return Object.freeze({
    inspectedCellCount: rowCount * columnCount,
    truncated,
    findings: Object.freeze(findings),
  });
}
