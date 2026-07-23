import {
  parseSpreadsheetDocument,
  type CellInput,
  type SparseCellInput,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../document';
import { InterchangeError } from './contracts';

export interface ImportedSheet {
  readonly name: string;
  readonly cells: readonly {
    readonly row: number;
    readonly column: number;
    readonly input: CellInput;
  }[];
}

export function buildDocument(sheets: readonly ImportedSheet[]): SpreadsheetDocument {
  const input: SpreadsheetDocumentInput = {
    schemaVersion: 2,
    id: 'interchange-document',
    workbook: {
      sheets: (sheets.length === 0 ? [{ name: 'Sheet 1', cells: [] }] : sheets).map(
        (sheet, index) => ({
          id: `interchange-sheet-${index + 1}`,
          name: sheet.name,
          cells: sheet.cells.map(
            ({ row, column, input: cellInput }): SparseCellInput => ({
              row,
              column,
              cell: { input: cellInput },
            }),
          ),
          merges: [],
        }),
      ),
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  };
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) {
    throw new InterchangeError('DOCUMENT_INVALID', 'Imported workbook is not a valid document');
  }
  return parsed.document;
}
