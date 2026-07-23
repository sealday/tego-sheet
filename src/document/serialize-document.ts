import type { SpreadsheetDocument } from './model/document';
import { parseSpreadsheetDocument } from './parse-document';

/**
 * Serializes a valid Workbook 2.0 snapshot to deterministic JSON bytes.
 *
 * @param document - Document snapshot to validate and serialize.
 * @returns Canonically ordered JSON text.
 * @throws TypeError when the supplied document is invalid.
 */
export function serializeSpreadsheetDocument(document: SpreadsheetDocument): string {
  const parsed = parseSpreadsheetDocument(document);
  if (!parsed.ok) {
    throw new TypeError('Cannot serialize an invalid spreadsheet document');
  }
  return JSON.stringify(parsed.document);
}
