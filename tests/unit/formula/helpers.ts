import { parseSpreadsheetDocument } from '../../../src/document';
import type { SpreadsheetDocument } from '../../../src/document';

export function formulaDocument(
  sheets: readonly {
    readonly id: string;
    readonly name: string;
    readonly cells: readonly {
      readonly row: number;
      readonly column: number;
      readonly input:
        | { readonly type: 'number'; readonly value: number }
        | { readonly type: 'string'; readonly value: string }
        | { readonly type: 'boolean'; readonly value: boolean }
        | { readonly type: 'formula'; readonly source: string };
    }[];
  }[],
  dateSystem: 'excel-1900' | 'excel-1904' = 'excel-1900',
): SpreadsheetDocument {
  const result = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'formula-document',
    workbook: {
      sheets: sheets.map((sheet) => ({
        ...sheet,
        cells: sheet.cells.map((cell) => ({
          row: cell.row,
          column: cell.column,
          cell: { input: cell.input },
        })),
        merges: [],
      })),
      styles: [],
      validations: [],
      settings: { dateSystem },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!result.ok) throw new Error(JSON.stringify(result.diagnostics));
  return result.document;
}
