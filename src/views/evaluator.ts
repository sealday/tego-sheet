import type { FormulaValue } from '../formula';
import type { FilterView, FilterViewPredicate } from './model';

/** Immutable inputs used to derive view visibility. */
export interface ApplyFilterViewInput {
  /** View definition to evaluate. */
  readonly view: FilterView;
  /** Source rows beginning at the view range start. */
  readonly rows: readonly (readonly FormulaValue[])[];
  /** Explicit comparison locale. */
  readonly locale: string;
  /** Evaluation resource limits. */
  readonly limits: {
    /** Maximum source rows accepted by one evaluation. */
    readonly maxRows: number;
  };
}

function scalar(value: FormulaValue | undefined): string | number | boolean | undefined {
  if (
    value === undefined ||
    value.type === 'blank' ||
    value.type === 'error' ||
    value.type === 'array'
  ) {
    return undefined;
  }
  return value.value;
}

function matches(
  value: string | number | boolean | undefined,
  filter: FilterViewPredicate,
  locale: string,
): boolean {
  if (value === undefined) return false;
  if (filter.operator === 'contains') {
    return String(value)
      .toLocaleLowerCase(locale)
      .includes(String(filter.value).toLocaleLowerCase(locale));
  }
  const comparison =
    typeof value === 'number' && typeof filter.value === 'number'
      ? value - filter.value
      : String(value).localeCompare(String(filter.value), locale);
  if (filter.operator === 'equal') return comparison === 0;
  if (filter.operator === 'notEqual') return comparison !== 0;
  if (filter.operator === 'greaterThan') return comparison > 0;
  if (filter.operator === 'greaterThanOrEqual') return comparison >= 0;
  if (filter.operator === 'lessThan') return comparison < 0;
  return comparison <= 0;
}

/** Derives hidden rows from a view without changing the source rows or document. */
export function applyFilterView(input: ApplyFilterViewInput): {
  /** Logical row indexes excluded by the active filters. */
  readonly hiddenRows: ReadonlySet<number>;
  /** Logical data-row indexes in the active sort order. */
  readonly rowOrder: readonly number[];
} {
  if (input.rows.length > input.limits.maxRows) {
    throw new RangeError('Filter view exceeds the configured row limit');
  }
  const hiddenRows = new Set<number>();
  const firstDataRow = input.view.range.start.row + 1;
  const lastDataRow = input.view.range.end.row;
  for (let row = firstDataRow; row <= lastDataRow; row += 1) {
    const values = input.rows[row - input.view.range.start.row];
    if (
      input.view.filters.some(
        (filter) =>
          !matches(
            scalar(values?.[filter.column - input.view.range.start.column]),
            filter,
            input.locale,
          ),
      )
    ) {
      hiddenRows.add(row);
    }
  }
  const collator = new Intl.Collator(input.locale, { numeric: true, sensitivity: 'base' });
  const compare = (
    left: string | number | boolean | undefined,
    right: string | number | boolean | undefined,
  ): number => {
    if (left === right) return 0;
    if (left === undefined) return 1;
    if (right === undefined) return -1;
    if (typeof left === 'number' && typeof right === 'number') return left - right;
    return collator.compare(String(left), String(right));
  };
  const rowOrder = Array.from(
    { length: Math.max(0, lastDataRow - firstDataRow + 1) },
    (_, index) => firstDataRow + index,
  ).sort((leftRow, rightRow) => {
    for (const sort of input.view.sorts) {
      const columnIndex = sort.column - input.view.range.start.column;
      const left = scalar(input.rows[leftRow - input.view.range.start.row]?.[columnIndex]);
      const right = scalar(input.rows[rightRow - input.view.range.start.row]?.[columnIndex]);
      const result = compare(left, right);
      if (result !== 0) return sort.direction === 'ascending' ? result : -result;
    }
    return leftRow - rightRow;
  });
  return { hiddenRows, rowOrder };
}
