import type { DocumentCellRange } from '../document';

/** One comparison applied to an absolute worksheet column. */
export interface FilterViewPredicate {
  readonly column: number;
  readonly operator:
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'greaterThanOrEqual'
    | 'lessThan'
    | 'lessThanOrEqual'
    | 'contains';
  readonly value: string | number | boolean;
}

/** One stable, serializable view definition. */
export interface FilterView {
  readonly id: string;
  readonly name: string;
  readonly range: DocumentCellRange;
  readonly sorts: readonly {
    readonly column: number;
    readonly direction: 'ascending' | 'descending';
  }[];
  readonly filters: readonly FilterViewPredicate[];
  readonly visibility: 'document' | 'session';
}
