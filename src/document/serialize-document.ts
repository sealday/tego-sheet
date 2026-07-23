import type { SpreadsheetDocument } from './model/document';
import { parseSpreadsheetDocument } from './parse-document';

export function serializeSpreadsheetDocument(document: SpreadsheetDocument): string {
  const parsed = parseSpreadsheetDocument(document);
  if (!parsed.ok) {
    throw new TypeError('Cannot serialize an invalid spreadsheet document');
  }
  return JSON.stringify(parsed.document);
}
