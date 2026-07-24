import type { StructuredTable } from './document';

/** True when a persisted table filter changes row visibility or ordering. */
export function hasActiveStructuredTableProjection(table: StructuredTable): boolean {
  return (
    table.filter !== undefined &&
    (table.filter.filters.some(({ operator }) => operator !== 'all') ||
      (table.filter.sort !== undefined && table.filter.sort !== null))
  );
}
