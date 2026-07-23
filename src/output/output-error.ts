import type { Diagnostic } from '../document';

/** Atomic output failure carrying one stable structured diagnostic. */
export class OutputAdapterError extends Error {
  /** Stable machine-readable failure code. */
  readonly code: string;
  /** Structured output diagnostic. */
  readonly diagnostic: Diagnostic;

  /** Creates an adapter failure from a structured diagnostic. */
  constructor(diagnostic: Diagnostic, options?: ErrorOptions) {
    super(diagnostic.message, options);
    this.name = 'OutputAdapterError';
    this.code = diagnostic.code;
    this.diagnostic = diagnostic;
  }
}

export function outputError(
  code: string,
  message: string,
  options: {
    readonly cause?: unknown;
    readonly details?: Diagnostic['details'];
    readonly location?: Diagnostic['location'];
    readonly stage?: Diagnostic['stage'];
  } = {},
): OutputAdapterError {
  return new OutputAdapterError(
    Object.freeze({
      code,
      severity: 'error',
      domain: 'output',
      stage: options.stage ?? 'serialize',
      message,
      ...(options.details === undefined ? {} : { details: options.details }),
      ...(options.location === undefined ? {} : { location: options.location }),
      ...(options.cause === undefined ? {} : { cause: options.cause }),
    }),
    options.cause === undefined ? undefined : { cause: options.cause },
  );
}

export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted === true) {
    throw outputError('RENDER_ABORTED', 'Output generation was aborted');
  }
}
