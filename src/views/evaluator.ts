import type { FormulaValue } from '../formula';
import type { FilterView, FilterViewPredicate } from './model';

export interface ApplyFilterViewInput {
  readonly view: FilterView;
  readonly rows: readonly (readonly FormulaValue[])[];
  readonly locale: string;
  readonly limits: { readonly maxRows: number };
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
  readonly hiddenRows: ReadonlySet<number>;
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
  return { hiddenRows };
}
