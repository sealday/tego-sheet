/** Stable planning failure. */
export class DataTransformError extends Error {
  /** Creates a machine-readable planning failure. */
  constructor(
    /** Stable transform-planning failure category. */
    readonly code:
      | 'TRANSFORM_ABORTED'
      | 'TRANSFORM_TOO_LARGE'
      | 'REPLACE_PATTERN_INVALID'
      | 'REPLACE_BUDGET_EXCEEDED'
      | 'TEXT_SPLIT_OVERFLOW'
      | 'FILL_SERIES_INVALID',
    message: string,
  ) {
    super(message);
    this.name = 'DataTransformError';
  }
}
