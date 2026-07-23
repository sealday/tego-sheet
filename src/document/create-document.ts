import type {
  CreateDocumentOptions,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from './model/document';
import { parseSpreadsheetDocument } from './parse-document';

/**
 * Creates a valid, deeply frozen empty Workbook 2.0 document.
 *
 * @param options - Optional stable IDs and initial workbook settings.
 * @returns A new immutable document snapshot.
 */
export function createSpreadsheetDocument(
  options: CreateDocumentOptions = {},
): SpreadsheetDocument {
  const id = options.id ?? globalThis.crypto.randomUUID();
  const sheetId = options.sheetId ?? globalThis.crypto.randomUUID();
  const input: SpreadsheetDocumentInput = {
    schemaVersion: 2,
    id,
    workbook: {
      sheets: [{ id: sheetId, name: options.sheetName ?? 'Sheet 1', cells: [], merges: [] }],
      styles: [],
      validations: [],
      settings: {
        dateSystem: options.dateSystem ?? 'excel-1900',
        ...(options.localeHint === undefined ? {} : { localeHint: options.localeHint }),
      },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  };
  const result = parseSpreadsheetDocument(input);
  if (!result.ok) throw new TypeError('Failed to create a valid spreadsheet document');
  return result.document;
}
