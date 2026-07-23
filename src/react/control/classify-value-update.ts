import { type TegoSheetError } from '../../core';
import { parseSpreadsheetDocument, type SpreadsheetDocument } from '../../document';
import type { PendingCheckpoint } from './pending-checkpoint';

export type ValueUpdate =
  | { readonly kind: 'same-reference' }
  | { readonly kind: 'acknowledge'; readonly through: number }
  | { readonly kind: 'rollback' }
  | { readonly kind: 'replace'; readonly document: SpreadsheetDocument }
  | { readonly kind: 'invalid'; readonly error: TegoSheetError };

export interface ValueClassificationState {
  readonly observedValue: unknown;
  readonly acknowledgedKey: string;
  readonly pending: readonly PendingCheckpoint[];
}

function invalidValue(cause: unknown): TegoSheetError {
  return {
    code: 'INVALID_DATA',
    message: 'Spreadsheet document is invalid',
    recoverable: true,
    cause,
  };
}

export function classifyValueUpdate(
  state: ValueClassificationState,
  value: SpreadsheetDocument,
): ValueUpdate {
  if (Object.is(value, state.observedValue)) return { kind: 'same-reference' };
  const parsed = parseSpreadsheetDocument(value);
  if (!parsed.ok) return { kind: 'invalid', error: invalidValue(parsed.diagnostics) };
  const key = JSON.stringify(parsed.document);
  for (let index = state.pending.length - 1; index >= 0; index -= 1) {
    if (state.pending[index]!.projectedKey === key) {
      return { kind: 'acknowledge', through: index };
    }
  }
  if (key === state.acknowledgedKey) return { kind: 'rollback' };
  return { kind: 'replace', document: parsed.document };
}
