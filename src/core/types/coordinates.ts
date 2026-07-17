declare const sheetIdBrand: unique symbol;

/** Opaque identifier assigned to a worksheet and returned through the public API. */
export type SheetId = string & {
  /** Prevents an arbitrary string from being used as a sheet identifier. */
  readonly [sheetIdBrand]: true;
};

/** A zero-based row and column coordinate within a sheet. */
export interface CellPoint {
  /** Zero-based row index. */
  readonly row: number;
  /** Zero-based column index. */
  readonly column: number;
}

/** A zero-based cell coordinate paired with its worksheet identifier. */
export interface CellAddress extends CellPoint {
  /** Worksheet containing the cell. */
  readonly sheet: SheetId;
}

/** An inclusive rectangular range normalized from top-left to bottom-right. */
export interface CellRange {
  /** Zero-based top-left corner, included in the range. */
  readonly start: CellPoint;
  /** Zero-based bottom-right corner, included in the range. */
  readonly end: CellPoint;
}

/** Current worksheet selection and its active cell. */
export interface Selection {
  /** Worksheet containing the selection. */
  readonly sheet: SheetId;
  /** Inclusive selected range. */
  readonly range: CellRange;
  /** Focused cell within `range`, used as the action anchor. */
  readonly active: CellPoint;
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function assertCoordinate(value: unknown, name: 'row' | 'column'): asserts value is number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new TypeError(`${name} must be a non-negative integer`);
  }
}

function assertSheetValue(value: unknown, name: 'sheet' | 'sheet id'): asserts value is SheetId {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new TypeError(`${name} must be a non-empty string`);
  }
}

function assertRangePoints(
  value: unknown,
): asserts value is { readonly start: CellPoint; readonly end: CellPoint } {
  if (!isObjectRecord(value)) {
    throw new TypeError('cell range must be an object');
  }
  if (!isObjectRecord(value.start)) {
    throw new TypeError('range start must be a cell point');
  }
  if (!isObjectRecord(value.end)) {
    throw new TypeError('range end must be a cell point');
  }

  assertCellPoint(value.start);
  assertCellPoint(value.end);
}

export function sheetId(value: string): SheetId {
  assertSheetValue(value, 'sheet id');
  return value;
}

export function assertCellPoint(value: unknown): asserts value is CellPoint {
  if (!isObjectRecord(value)) {
    throw new TypeError('cell point must be an object');
  }

  assertCoordinate(value.row, 'row');
  assertCoordinate(value.column, 'column');
}

export function assertCellAddress(value: unknown): asserts value is CellAddress {
  if (!isObjectRecord(value)) {
    throw new TypeError('cell address must be an object');
  }

  assertSheetValue(value.sheet, 'sheet');
  assertCoordinate(value.row, 'row');
  assertCoordinate(value.column, 'column');
}

export function assertCellRange(value: unknown): asserts value is CellRange {
  assertRangePoints(value);

  if (value.start.row > value.end.row || value.start.column > value.end.column) {
    throw new RangeError('cell range must be normalized from top-left to bottom-right');
  }
}

export function normalizeCellRange(value: {
  readonly start: CellPoint;
  readonly end: CellPoint;
}): CellRange {
  assertRangePoints(value);

  return {
    start: {
      row: Math.min(value.start.row, value.end.row),
      column: Math.min(value.start.column, value.end.column),
    },
    end: {
      row: Math.max(value.start.row, value.end.row),
      column: Math.max(value.start.column, value.end.column),
    },
  };
}
