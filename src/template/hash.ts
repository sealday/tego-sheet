import type { SpreadsheetDocument } from '../document';

/** Computes a stable non-cryptographic source identity from canonical schema bytes. */
export function hashSpreadsheetDocument(document: SpreadsheetDocument): string {
  const canonicalize = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(canonicalize);
    if (value !== null && typeof value === 'object') {
      return Object.fromEntries(
        Object.entries(value as Readonly<Record<string, unknown>>)
          .filter(([, child]) => child !== undefined)
          .sort(([left], [right]) => left.localeCompare(right))
          .map(([key, child]) => [key, canonicalize(child)]),
      );
    }
    return value;
  };
  const source = JSON.stringify(canonicalize(document));
  let first = 0xcbf29ce4;
  let second = 0x84222325;
  for (let index = 0; index < source.length; index += 1) {
    const code = source.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193) >>> 0;
    second = Math.imul(second ^ (code + index), 0x01000193) >>> 0;
  }
  return `${first.toString(16).padStart(8, '0')}${second.toString(16).padStart(8, '0')}`;
}
