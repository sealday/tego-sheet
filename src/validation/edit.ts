import type { DocumentController, DocumentTransactionResult } from '../document-controller';
import type { SheetId } from '../core';
import type { ChangeSource } from '../core';
import type { DocumentCellAddress } from '../document';
import type { ValidationEngineOptions } from './engine';
import type { ValidationResult } from './model';
import { executeValidatedTransaction } from './transaction';

/** Inputs for validating and atomically committing one document-owned cell edit. */
export interface ValidatedCellEditRequest {
  /** Document controller that owns mutation, validation rules, and history. */
  readonly controller: DocumentController;
  /** Cell receiving the submitted text. */
  readonly address: DocumentCellAddress;
  /** User-facing cell text submitted for validation and commit. */
  readonly text: string;
  /** Restricted resolver/formula capabilities for document-owned rules. */
  readonly validation?: ValidationEngineOptions;
  /** Explicit confirmation gate for warning-mode rules. */
  readonly confirmWarning?: (result: ValidationResult) => boolean | Promise<boolean>;
  /** Interaction surface attributed to an accepted edit transaction. */
  readonly source?: ChangeSource;
  /** Optional cancellation signal shared with async validation. */
  readonly signal?: AbortSignal;
  /** Optional lifecycle/race gate checked immediately before mutation. */
  readonly canCommit?: () => boolean;
}

/** Single transaction result used to commit a document-owned validated edit. */
export type ValidatedCellEditResult = DocumentTransactionResult;

/** Resolves the current document rule, validates the submitted text, and commits atomically. */
export async function executeValidatedCellEdit(
  input: ValidatedCellEditRequest,
): Promise<ValidatedCellEditResult> {
  const baseRevision = input.controller.getSnapshot().revision;
  return executeValidatedTransaction({
    controller: input.controller,
    transaction: {
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
              sheet: input.address.sheetId as string as SheetId,
              row: input.address.row,
              column: input.address.column,
            },
            text: input.text,
          },
        },
      ],
    },
    ...(input.validation === undefined ? {} : { validation: input.validation }),
    ...(input.confirmWarning === undefined ? {} : { confirmWarning: input.confirmWarning }),
    ...(input.source === undefined ? {} : { options: { source: input.source } }),
    ...(input.signal === undefined ? {} : { signal: input.signal }),
    ...(input.canCommit === undefined ? {} : { canCommit: input.canCommit }),
  });
}
