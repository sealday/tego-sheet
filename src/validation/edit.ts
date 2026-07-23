import type { DocumentController, DocumentTransactionResult } from '../document-controller';
import type { SheetId } from '../core';
import type { ValidationEngine } from './engine';
import type { ValidationRequest, ValidationResult } from './model';

/** Inputs for validating and atomically committing one cell edit. */
export interface ValidatedCellEditRequest {
  /** Document controller that owns mutation and history. */
  readonly controller: DocumentController;
  /** Validation engine used before mutation. */
  readonly engine: ValidationEngine;
  /** Typed validation candidate. */
  readonly request: ValidationRequest;
  /** User-facing cell text submitted after validation. */
  readonly text: string;
  /** Explicit confirmation gate for warning-mode rules. */
  readonly confirmWarning?: (result: ValidationResult) => boolean | Promise<boolean>;
}

/** Validation result or the single transaction result used to commit an accepted edit. */
export type ValidatedCellEditResult = ValidationResult | DocumentTransactionResult;

/** Validates an edit and commits it through exactly one document transaction. */
export async function executeValidatedCellEdit(
  input: ValidatedCellEditRequest,
): Promise<ValidatedCellEditResult> {
  const validation = await input.engine.validate(input.request);
  if (validation.status === 'rejected' || validation.status === 'error') return validation;
  if (
    validation.status === 'warning' &&
    (input.confirmWarning === undefined || !(await input.confirmWarning(validation)))
  ) {
    return validation;
  }
  const revision = input.controller.getSnapshot().revision;
  return input.controller.transact({
    schemaVersion: 1,
    id: `validated-edit-${revision}`,
    baseRevision: revision,
    commands: [
      {
        schemaVersion: 1,
        id: 'set-validated-cell',
        command: {
          type: 'set-cell-text',
          address: {
            sheet: input.request.address.sheetId as string as SheetId,
            row: input.request.address.row,
            column: input.request.address.column,
          },
          text: input.text,
        },
      },
    ],
  });
}
