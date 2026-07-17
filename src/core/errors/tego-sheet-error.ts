/** Stable category for a public spreadsheet error. */
export type TegoSheetErrorCode =
  | 'INVALID_DATA'
  | 'INVALID_COMMAND'
  | 'CLIPBOARD_DENIED'
  | 'PRINT_FAILED'
  | 'RENDER_FAILED';

/**
 * Structured error payload delivered to `onError` for an operation the component handled.
 * Thrown contract and imperative API failures use `TegoSheetException` instead.
 */
export interface TegoSheetError {
  /** Machine-readable error category. */
  readonly code: TegoSheetErrorCode;
  /** Human-readable description of the failure. */
  readonly message: string;
  /** Whether the component can continue operating after the failure. */
  readonly recoverable: boolean;
  /** Original failure when one is available. */
  readonly cause?: unknown;
}
