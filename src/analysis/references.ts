import type { DocumentCellRange } from '../document';

const maximumRow = 1_048_575;
const maximumColumn = 16_383;
const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/;

/** Immutable, stable worksheet range consumed by analysis objects. */
export type AnalysisRangeReference = DocumentCellRange;

/** Immutable, stable worksheet address consumed by analysis objects. */
export interface AnalysisCellReference {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

function coordinate(value: number, maximum: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 0 || value > maximum) {
    throw new RangeError(`${label} is outside the worksheet bounds`);
  }
  return value;
}

export function snapshotRange(
  reference: AnalysisRangeReference,
  label: string,
): AnalysisRangeReference {
  if (
    reference === null ||
    typeof reference !== 'object' ||
    !identifierPattern.test(reference.sheetId)
  ) {
    throw new TypeError(`${label} is invalid`);
  }
  const start = Object.freeze({
    row: coordinate(reference.start.row, maximumRow, `${label} start row`),
    column: coordinate(reference.start.column, maximumColumn, `${label} start column`),
  });
  const end = Object.freeze({
    row: coordinate(reference.end.row, maximumRow, `${label} end row`),
    column: coordinate(reference.end.column, maximumColumn, `${label} end column`),
  });
  if (start.row > end.row || start.column > end.column) {
    throw new RangeError(`${label} must be normalized`);
  }
  return Object.freeze({ sheetId: reference.sheetId, start, end });
}

export function snapshotAddress(
  reference: AnalysisCellReference,
  label: string,
): AnalysisCellReference {
  if (
    reference === null ||
    typeof reference !== 'object' ||
    !identifierPattern.test(reference.sheetId)
  ) {
    throw new TypeError(`${label} is invalid`);
  }
  return Object.freeze({
    sheetId: reference.sheetId,
    row: coordinate(reference.row, maximumRow, `${label} row`),
    column: coordinate(reference.column, maximumColumn, `${label} column`),
  });
}

export function rangeCellCount(reference: AnalysisRangeReference): number {
  return (
    (reference.end.row - reference.start.row + 1) *
    (reference.end.column - reference.start.column + 1)
  );
}

export function rangesIntersect(
  left: AnalysisRangeReference,
  right: AnalysisRangeReference,
): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
}

export function checkedIdentifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

export function checkedPositiveLimit(
  value: number | undefined,
  fallback: number,
  label: string,
): number {
  const result = value ?? fallback;
  if (!Number.isSafeInteger(result) || result < 1) {
    throw new RangeError(`${label} must be a positive safe integer`);
  }
  return result;
}
