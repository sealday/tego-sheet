import type { DocumentCellRange, SparseCell } from '../document';
import { DataTransformError } from './errors';

const CHUNK_SIZE = 256;

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new DataTransformError('TRANSFORM_ABORTED', 'Data operation was aborted');
  }
}

function compareCell(
  cell: Pick<SparseCell, 'row' | 'column'>,
  point: { readonly row: number; readonly column: number },
): number {
  return cell.row - point.row || cell.column - point.column;
}

/** Yields to the host event loop at bounded intervals and rechecks cancellation. */
export async function yieldForCancellation(
  completed: number,
  signal: AbortSignal | undefined,
): Promise<void> {
  throwIfAborted(signal);
  if (completed === 0 || completed % CHUNK_SIZE !== 0) return;
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  throwIfAborted(signal);
}

/** Indexes only the canonical sparse-cell slice intersecting one bounded range. */
export async function indexSparseRange(
  cells: readonly SparseCell[],
  range: DocumentCellRange,
  signal: AbortSignal | undefined,
): Promise<ReadonlyMap<string, SparseCell>> {
  let low = 0;
  let high = cells.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (compareCell(cells[middle] as SparseCell, range.start) < 0) low = middle + 1;
    else high = middle;
  }
  const indexed = new Map<string, SparseCell>();
  let scanned = 0;
  for (let index = low; index < cells.length; index += 1) {
    const cell = cells[index] as SparseCell;
    if (cell.row > range.end.row) break;
    if (cell.row === range.end.row && cell.column > range.end.column) break;
    if (cell.column >= range.start.column && cell.column <= range.end.column) {
      indexed.set(`${cell.row}:${cell.column}`, cell);
    }
    scanned += 1;
    await yieldForCancellation(scanned, signal);
  }
  return indexed;
}
