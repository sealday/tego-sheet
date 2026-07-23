import type { DocumentCellRange } from '../document';

/** One comparison applied to an absolute worksheet column. */
export interface FilterViewPredicate {
  /** Absolute worksheet column. */
  readonly column: number;
  /** Scalar comparison operator. */
  readonly operator:
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'greaterThanOrEqual'
    | 'lessThan'
    | 'lessThanOrEqual'
    | 'contains';
  /** Fixed comparison value. */
  readonly value: string | number | boolean;
}

/** One stable, serializable view definition. */
export interface FilterView {
  /** Stable view identifier. */
  readonly id: string;
  /** User-visible view name. */
  readonly name: string;
  /** Header-inclusive source range. */
  readonly range: DocumentCellRange;
  /** Ordered sort definitions. */
  readonly sorts: readonly {
    readonly column: number;
    readonly direction: 'ascending' | 'descending';
  }[];
  /** Predicates combined with logical AND. */
  readonly filters: readonly FilterViewPredicate[];
  /** Whether the definition is persistent or session-owned. */
  readonly visibility: 'document' | 'session';
}
