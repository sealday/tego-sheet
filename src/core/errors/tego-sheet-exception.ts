import type { TegoSheetError, TegoSheetErrorCode } from './tego-sheet-error';

/** Error thrown when a public call or component contract cannot complete with a {@link TegoSheetError}. */
export class TegoSheetException extends Error {
  /** Identifies this error as a TegoSheet exception. */
  override readonly name = 'TegoSheetException';
  /** Original failure when one is available. */
  override readonly cause?: unknown;
  /** Complete structured error payload. */
  readonly error: TegoSheetError;
  /** Machine-readable category copied from `error`. */
  readonly code: TegoSheetErrorCode;
  /** Whether the mounted component can continue operating after the failure. */
  readonly recoverable: boolean;

  /** Creates a throwable wrapper around a structured spreadsheet error. */
  constructor(error: TegoSheetError) {
    super(error.message, { cause: error.cause });
    this.cause = error.cause;
    this.error = error;
    this.code = error.code;
    this.recoverable = error.recoverable;
  }
}
