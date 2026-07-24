import type { DisplayRect } from '../print';
import type { ObjectAnchor, ObjectCoordinateTransform, ObjectGeometry } from './model';

/** Resolves a stable worksheet anchor to device-independent geometry. */
export function resolveObjectAnchor(anchor: ObjectAnchor, geometry: ObjectGeometry): DisplayRect {
  if (anchor.type === 'absolute') return { ...anchor.rect };
  if (anchor.type === 'one-cell') {
    return {
      x: geometry.columnOffset(anchor.cell.column) + anchor.offset.x,
      y: geometry.rowOffset(anchor.cell.row) + anchor.offset.y,
      width: anchor.size.width,
      height: anchor.size.height,
    };
  }
  const x = geometry.columnOffset(anchor.from.column) + anchor.from.offset.x;
  const y = geometry.rowOffset(anchor.from.row) + anchor.from.offset.y;
  const right = geometry.columnOffset(anchor.to.column) + anchor.to.offset.x;
  const bottom = geometry.rowOffset(anchor.to.row) + anchor.to.offset.y;
  return { x, y, width: Math.max(0, right - x), height: Math.max(0, bottom - y) };
}

function transformIndex(
  value: number,
  operation: ObjectCoordinateTransform,
  axis: 'row' | 'column',
): number {
  const matchesAxis = operation.type.endsWith(axis);
  if (!matchesAxis) return value;
  if (operation.type.startsWith('insert')) {
    return value >= operation.index ? value + operation.count : value;
  }
  if (value < operation.index) return value;
  if (value >= operation.index + operation.count) return value - operation.count;
  return operation.index;
}

/** Applies the same row/column coordinate transform used by sheet commands. */
export function transformObjectAnchor(
  anchor: ObjectAnchor,
  operation: ObjectCoordinateTransform,
): ObjectAnchor {
  if (anchor.type === 'absolute') return anchor;
  if (anchor.type === 'one-cell') {
    if (anchor.cell.sheetId !== operation.sheetId) return anchor;
    return {
      ...anchor,
      cell: {
        ...anchor.cell,
        row: transformIndex(anchor.cell.row, operation, 'row'),
        column: transformIndex(anchor.cell.column, operation, 'column'),
      },
    };
  }
  if (anchor.from.sheetId !== operation.sheetId || anchor.to.sheetId !== operation.sheetId) {
    return anchor;
  }
  const transformed: Extract<ObjectAnchor, { type: 'two-cell' }> = {
    ...anchor,
    from: {
      ...anchor.from,
      row: transformIndex(anchor.from.row, operation, 'row'),
      column: transformIndex(anchor.from.column, operation, 'column'),
    },
    to: {
      ...anchor.to,
      row: transformIndex(anchor.to.row, operation, 'row'),
      column: transformIndex(anchor.to.column, operation, 'column'),
    },
  };
  if (operation.type === 'delete-row' && transformed.from.row === transformed.to.row) {
    return {
      ...transformed,
      from: { ...transformed.from, offset: { ...transformed.from.offset, y: 0 } },
      to: { ...transformed.to, offset: { ...transformed.to.offset, y: 0 } },
    };
  }
  if (operation.type === 'delete-column' && transformed.from.column === transformed.to.column) {
    return {
      ...transformed,
      from: { ...transformed.from, offset: { ...transformed.from.offset, x: 0 } },
      to: { ...transformed.to, offset: { ...transformed.to.offset, x: 0 } },
    };
  }
  return transformed;
}
