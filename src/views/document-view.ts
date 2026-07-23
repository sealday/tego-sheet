import type { CellInput, SpreadsheetDocument } from '../document';
import { formulaAddressKey, type FormulaValue } from '../formula';
import type { FilterView } from './model';
import { applyFilterView } from './evaluator';

/** Immutable document inputs used to derive one renderer-ready saved view. */
export interface ApplyDocumentFilterViewInput {
  /** Source document snapshot. */
  readonly document: SpreadsheetDocument;
  /** Formula results for the same document revision. */
  readonly formulaValues?: ReadonlyMap<string, FormulaValue>;
  /** Selected document or session view. */
  readonly view: FilterView;
  /** Explicit comparison locale. */
  readonly locale: string;
  /** Evaluation resource limits. */
  readonly limits: { readonly maxRows: number };
}

function inputValue(input: CellInput | undefined): FormulaValue {
  if (input === undefined || input.type === 'blank' || input.type === 'formula') {
    return { type: 'blank' };
  }
  if (input.type === 'custom') return { type: 'blank' };
  if (input.type === 'number') return { type: 'number', value: input.value };
  if (input.type === 'boolean') return { type: 'boolean', value: input.value };
  return { type: 'string', value: input.value };
}

/** Projects document inputs and formula results into derived visibility and row order. */
export function applyDocumentFilterView(input: ApplyDocumentFilterViewInput): {
  readonly hiddenRows: ReadonlySet<number>;
  readonly rowOrder: readonly number[];
} {
  const sheet = input.document.workbook.sheets.find(({ id }) => id === input.view.range.sheetId);
  if (sheet === undefined) throw new RangeError(`Unknown view sheet: ${input.view.range.sheetId}`);
  const cells = new Map(sheet.cells.map((entry) => [`${entry.row}:${entry.column}`, entry.cell]));
  const rows: FormulaValue[][] = [];
  for (let row = input.view.range.start.row; row <= input.view.range.end.row; row += 1) {
    const values: FormulaValue[] = [];
    for (
      let column = input.view.range.start.column;
      column <= input.view.range.end.column;
      column += 1
    ) {
      const cell = cells.get(`${row}:${column}`);
      values.push(
        cell?.input.type === 'formula'
          ? (input.formulaValues?.get(formulaAddressKey({ sheetId: sheet.id, row, column })) ?? {
              type: 'blank',
            })
          : inputValue(cell?.input),
      );
    }
    rows.push(values);
  }
  return applyFilterView({
    view: input.view,
    rows,
    locale: input.locale,
    limits: input.limits,
  });
}
