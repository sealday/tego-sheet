import type { DocumentController, DocumentTransactionResult } from '../document-controller';
import type { SheetId } from '../core';
import type { ChangeSource } from '../core';
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
  /** Interaction surface attributed to an accepted edit transaction. */
  readonly source?: ChangeSource;
  /** Optional lifecycle/race gate checked immediately before mutation. */
  readonly canCommit?: () => boolean;
}

/** Validation result or the single transaction result used to commit an accepted edit. */
export type ValidatedCellEditResult = ValidationResult | DocumentTransactionResult;

/** Validates an edit and commits it through exactly one document transaction. */
export async function executeValidatedCellEdit(
  input: ValidatedCellEditRequest,
): Promise<ValidatedCellEditResult> {
  const baseRevision = input.controller.getSnapshot().revision;
  const validation = await input.engine.validate(input.request);
  if (validation.status === 'rejected' || validation.status === 'error') return validation;
  if (
    validation.status === 'warning' &&
    (input.confirmWarning === undefined || !(await input.confirmWarning(validation)))
  ) {
    return validation;
  }
  if (input.canCommit?.() === false) {
    return {
      status: 'error',
      code: 'VALIDATION_SOURCE_ERROR',
      diagnostics: [{ code: 'VALIDATION_EDIT_CANCELLED', ruleId: input.request.rule.id }],
    };
  }
  return input.controller.transact(
    {
      schemaVersion: 1,
      id: `validated-edit-${baseRevision}`,
      baseRevision,
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
    },
    input.source === undefined ? undefined : { source: input.source },
  );
}
