import type { Cell } from './document';

/** A non-empty cell stored at an explicit zero-based coordinate. */
export interface SparseCell {
  /** Zero-based row index. */
  readonly row: number;
  /** Zero-based column index. */
  readonly column: number;
  /** Persistent cell data. */
  readonly cell: Cell;
}

export interface SparseCellInput {
  row: number;
  column: number;
  cell: CellInputRecord;
}

export interface CellInputRecord {
  input: import('./document').CellInput;
  styleId?: string;
  validationId?: string;
  resourceId?: string;
  templateId?: string;
  metadata?: import('../../core/types/json').JsonValue;
}

export function compareSparseCells(
  left: Pick<SparseCell, 'row' | 'column'>,
  right: Pick<SparseCell, 'row' | 'column'>,
): number {
  return left.row - right.row || left.column - right.column;
}
