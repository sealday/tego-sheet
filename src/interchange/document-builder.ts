import {
  parseSpreadsheetDocument,
  type CellInput,
  type ConditionalFormat,
  type JsonValue,
  type SheetFilter,
  type SheetRange,
  type SparseCellInput,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
  type WorksheetVisibility,
} from '../document';
import { InterchangeError } from './contracts';

export interface ImportedSheet {
  readonly id?: string;
  readonly name: string;
  readonly cells: readonly {
    readonly row: number;
    readonly column: number;
    readonly input: CellInput;
    readonly styleId?: string;
    readonly validationId?: string;
  }[];
  readonly merges?: readonly SheetRange[];
  readonly filter?: SheetFilter;
  readonly conditionalFormatting?: readonly ConditionalFormat[];
  readonly visibility?: WorksheetVisibility;
}

export interface ImportedWorkbookMetadata {
  readonly styles?: readonly { readonly id: string; readonly value: JsonValue }[];
  readonly validations?: readonly { readonly id: string; readonly value: JsonValue }[];
  readonly templates?: readonly SpreadsheetDocumentInput['templates'][number][];
}

export function buildDocument(
  sheets: readonly ImportedSheet[],
  metadata: ImportedWorkbookMetadata = {},
): SpreadsheetDocument {
  const input: SpreadsheetDocumentInput = {
    schemaVersion: 2,
    id: 'interchange-document',
    workbook: {
      sheets: (sheets.length === 0 ? [{ name: 'Sheet 1', cells: [] }] : sheets).map(
        (sheet, index) => ({
          id: sheet.id ?? `interchange-sheet-${index + 1}`,
          name: sheet.name,
          cells: sheet.cells.map(
            ({ row, column, input: cellInput, styleId, validationId }): SparseCellInput => ({
              row,
              column,
              cell: {
                input: cellInput,
                ...(styleId === undefined ? {} : { styleId }),
                ...(validationId === undefined ? {} : { validationId }),
              },
            }),
          ),
          merges:
            sheet.merges?.map((range) => ({
              start: { ...range.start },
              end: { ...range.end },
            })) ?? [],
          ...(sheet.filter === undefined
            ? {}
            : {
                filter: {
                  ...(sheet.filter.range === undefined
                    ? {}
                    : {
                        range: {
                          start: { ...sheet.filter.range.start },
                          end: { ...sheet.filter.range.end },
                        },
                      }),
                  filters: sheet.filter.filters.map((filter) => ({
                    column: filter.column,
                    operator: filter.operator,
                    values: [...filter.values],
                  })),
                  ...(sheet.filter.sort === undefined
                    ? {}
                    : {
                        sort: sheet.filter.sort === null ? null : { ...sheet.filter.sort },
                      }),
                },
              }),
          ...(sheet.visibility === undefined ? {} : { visibility: sheet.visibility }),
          ...(sheet.conditionalFormatting === undefined
            ? {}
            : { conditionalFormatting: [...sheet.conditionalFormatting] }),
        }),
      ),
      styles: metadata.styles?.map((entry) => ({ id: entry.id, value: entry.value })) ?? [],
      validations:
        metadata.validations?.map((entry) => ({ id: entry.id, value: entry.value })) ?? [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates:
      metadata.templates?.map((template) => ({
        id: template.id,
        name: template.name,
        bindings: [...template.bindings],
        printProfiles: [...template.printProfiles],
      })) ?? [],
    resources: { items: [] },
    extensions: {},
  };
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) {
    throw new InterchangeError('DOCUMENT_INVALID', 'Imported workbook is not a valid document');
  }
  return parsed.document;
}
