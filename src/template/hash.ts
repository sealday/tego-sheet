import type { SpreadsheetDocument } from '../document';
import { serializeSpreadsheetDocument } from '../document';

/** Computes a stable non-cryptographic source identity from canonical schema bytes. */
export function hashSpreadsheetDocument(document: SpreadsheetDocument): string {
  const source = serializeSpreadsheetDocument(document);
  let first = 0xcbf29ce4;
  let second = 0x84222325;
  for (let index = 0; index < source.length; index += 1) {
    const code = source.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193) >>> 0;
    second = Math.imul(second ^ (code + index), 0x01000193) >>> 0;
  }
  return `${first.toString(16).padStart(8, '0')}${second.toString(16).padStart(8, '0')}`;
}
