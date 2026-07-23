import { createFormulaEvaluationBudget, evaluateCell } from '../core/formulas/evaluator';
import { formatValue, renderFormulaValue } from '../core/formulas/rendered-value';
import { getCellData } from '../core/model/cells';
import { selectCellStyle } from '../core/selectors/style';
import type { CellPoint } from '../core/types/coordinates';
import type { CellStyle, SheetData } from '../core/types/workbook';
import type { DocumentSheetId } from '../document';
import type { FormulaValue } from '../formula';
import type { CellPresentation, ResolvedStyle } from './cell-presentation';

/** Read-only compatibility resolver for the legacy Canvas projection. */
export interface LegacyPresentationResolver {
  resolve(point: CellPoint, target: 'screen' | 'print'): CellPresentation;
}

function source(sheet: Readonly<SheetData>, point: CellPoint): string | number | boolean | null {
  return getCellData(sheet, point.row, point.column)?.text ?? null;
}

function formulaValue(value: string | number | boolean): FormulaValue {
  if (typeof value === 'number') return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  if (value === '#ERROR!') return { type: 'error', value: '#VALUE!' };
  return { type: 'string', value };
}

/** Converts legacy style projection to the shared read-only style contract. */
export function resolveLegacyStyle(style: CellStyle): ResolvedStyle {
  return Object.freeze({
    color: style.color ?? '#0a0a0a',
    backgroundColor: style.bgcolor ?? '#ffffff',
    fontFamily: style.font?.name ?? 'Arial',
    fontSize: style.font?.size ?? 10,
    bold: style.font?.bold ?? false,
    italic: style.font?.italic ?? false,
    horizontalAlign: style.align ?? 'left',
    verticalAlign: style.valign ?? 'middle',
    wrap: style.textwrap === true,
    ...(style.format === undefined ? {} : { numberFormat: style.format }),
    ...(style.underline === undefined ? {} : { underline: style.underline }),
    ...(style.strike === undefined ? {} : { strike: style.strike }),
    ...(style.border === undefined ? {} : { border: style.border }),
  });
}

/** Creates one batch-scoped legacy resolver with a shared evaluation budget. */
export function createLegacyPresentationResolver(
  sheet: Readonly<SheetData>,
  defaultStyle: CellStyle,
  sheetId = 'legacy-sheet' as DocumentSheetId,
): LegacyPresentationResolver {
  const budget = createFormulaEvaluationBudget(250_000);
  const resolver: LegacyPresentationResolver = {
    resolve(point: CellPoint, target: 'screen' | 'print') {
      const cell = getCellData(sheet, point.row, point.column);
      const legacyStyle = selectCellStyle(sheet, point.row, point.column, defaultStyle);
      const printable = cell?.printable !== false;
      let rendered: string | number | boolean;
      if (target === 'print' && !printable) rendered = '';
      else if (
        typeof cell?.text === 'string' &&
        cell.text.startsWith('=') &&
        cell.value !== undefined
      ) {
        rendered = cell.value as string | number | boolean;
      } else {
        try {
          rendered = evaluateCell(point, (candidate) => source(sheet, candidate), budget);
        } catch {
          rendered = '#ERROR!';
        }
      }
      const formattedText = formatValue(legacyStyle.format, renderFormulaValue(rendered));
      const readOnly = cell?.editable === false;
      const invalid = rendered === '#ERROR!';
      return Object.freeze({
        address: Object.freeze({ sheetId, ...point }),
        value: Object.freeze(formulaValue(rendered)),
        formattedText,
        style: resolveLegacyStyle(legacyStyle),
        validation: Object.freeze(
          invalid
            ? { status: 'error' as const, message: 'Formula evaluation failed' }
            : { status: 'valid' as const },
        ),
        annotations: Object.freeze([]),
        visibility: Object.freeze({ hidden: false, printable }),
        accessibility: Object.freeze({
          label: formattedText,
          readOnly,
          invalid,
          ...(invalid ? { description: 'Formula evaluation failed' } : {}),
        }),
      });
    },
  };
  return Object.freeze(resolver);
}
