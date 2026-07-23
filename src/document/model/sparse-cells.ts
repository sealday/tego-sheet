import type { Cell } from './document';

export interface SparseCell {
  readonly row: number;
  readonly column: number;
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
