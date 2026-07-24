import {
  commitValidatedDocumentTransaction,
  previewValidatedDocumentTransaction,
  type DocumentController,
  type DocumentTransactionEnvelope,
  type DocumentTransactionOptions,
  type DocumentTransactionResult,
} from '../document-controller';
import type { SheetId } from '../core';
import { internalPasteRange } from '../core/operations/clipboard';
import {
  snapshotSerializableTransaction,
  type SerializableTransactionEnvelope,
} from '../core/controller/spreadsheet-document-controller';
import type { SpreadsheetDocument } from '../document';
import { createValidationEngine, type ValidationEngineOptions } from './engine';
import type { ValidationRequest, ValidationResult } from './model';
import { resolveDocumentValidation } from './document-rule';

interface ValidationCandidate {
  readonly sheet: SheetId;
  readonly row: number;
  readonly column: number;
  readonly text: string;
}

const MAX_VALIDATION_CANDIDATES = 100_000;

/** Inputs for validating every content edit before one atomic transaction commit. */
export interface ValidatedTransactionRequest {
  /** Document controller that owns mutation and history. */
  readonly controller: DocumentController;
  /** Restricted resolver/formula capabilities used by an internally owned validation engine. */
  readonly validation?: ValidationEngineOptions;
  /** Atomic content transaction to validate and commit. */
  readonly transaction: DocumentTransactionEnvelope;
  /** Optional transaction source and permission gate. */
  readonly options?: DocumentTransactionOptions;
  /** Optional cancellation signal shared by all validation requests. */
  readonly signal?: AbortSignal;
  /** Explicit confirmation gate for warning-mode rules. */
  readonly confirmWarning?: (result: ValidationResult) => boolean | Promise<boolean>;
  /** Optional lifecycle/race gate checked immediately before mutation. */
  readonly canCommit?: () => boolean;
  /** Batch-wide resource bounds in addition to per-resolver engine limits. */
  readonly limits?: {
    /** Maximum validations evaluated concurrently. */
    readonly maxConcurrency?: number;
    /** Hard deadline for the complete validation batch. */
    readonly totalTimeoutMs?: number;
  };
}

function rejected(code: string, message: string): DocumentTransactionResult {
  return { status: 'rejected', code, message };
}

function aborted(signal: AbortSignal | undefined): boolean {
  return signal?.aborted === true;
}

function validPositiveLimit(value: number): boolean {
  return Number.isFinite(value) && Number.isInteger(value) && value > 0;
}

async function beforeDeadline<Value>(value: Promise<Value>, signal: AbortSignal): Promise<Value> {
  if (signal.aborted) throw signal.reason ?? new Error('Validation aborted');
  return new Promise<Value>((resolve, reject) => {
    const onAbort = () => {
      signal.removeEventListener('abort', onAbort);
      reject(signal.reason ?? new Error('Validation aborted'));
    };
    signal.addEventListener('abort', onAbort, { once: true });
    value.then(
      (result) => {
        signal.removeEventListener('abort', onAbort);
        resolve(result);
      },
      (error: unknown) => {
        signal.removeEventListener('abort', onAbort);
        reject(error);
      },
    );
  });
}

function candidateText(
  document: SpreadsheetDocument,
  candidate: Pick<ValidationCandidate, 'sheet' | 'row' | 'column'>,
): string {
  const input = document.workbook.sheets
    .find((sheet) => sheet.id === (candidate.sheet as string))
    ?.cells.find((cell) => cell.row === candidate.row && cell.column === candidate.column)
    ?.cell.input;
  if (input === undefined || input.type === 'blank') return '';
  if (input.type === 'string') return input.value;
  if (input.type === 'number' || input.type === 'boolean') return String(input.value);
  if (input.type === 'formula') return input.source;
  return typeof input.value === 'string' ? input.value : JSON.stringify(input.value);
}

function candidates(
  transaction: DocumentTransactionEnvelope,
):
  | { readonly ok: true; readonly values: readonly ValidationCandidate[] }
  | { readonly ok: false; readonly result: DocumentTransactionResult } {
  const output: ValidationCandidate[] = [];
  const reserve = (count: number): boolean =>
    Number.isSafeInteger(count) && count >= 0 && output.length + count <= MAX_VALIDATION_CANDIDATES;
  for (const envelope of transaction.commands) {
    const command = envelope.command;
    if (command.type === 'set-cell-text') {
      if (!reserve(1)) {
        return {
          ok: false,
          result: rejected('TRANSACTION_LIMIT_EXCEEDED', 'Validation candidate limit exceeded'),
        };
      }
      output.push({ ...command.address, text: command.text });
      continue;
    }
    if (command.type === 'clear-contents') {
      const { start, end } = command.selection.range;
      const rows = end.row - start.row + 1;
      const columns = end.column - start.column + 1;
      const count = rows * columns;
      if (!reserve(count)) {
        return {
          ok: false,
          result: rejected('TRANSACTION_LIMIT_EXCEEDED', 'Validation candidate limit exceeded'),
        };
      }
      for (let row = start.row; row <= end.row; row += 1) {
        for (let column = start.column; column <= end.column; column += 1) {
          output.push({ sheet: command.selection.sheet, row, column, text: '' });
        }
      }
      continue;
    }
    if (command.type === 'paste-external') {
      const count = command.values.reduce((total, row) => total + row.length, 0);
      if (!reserve(count)) {
        return {
          ok: false,
          result: rejected('TRANSACTION_LIMIT_EXCEEDED', 'Validation candidate limit exceeded'),
        };
      }
      const start = command.target.range.start;
      for (let row = 0; row < command.values.length; row += 1) {
        const values = command.values[row]!;
        for (let column = 0; column < values.length; column += 1) {
          output.push({
            sheet: command.target.sheet,
            row: start.row + row,
            column: start.column + column,
            text: values[column]!,
          });
        }
      }
      continue;
    }
    if (
      (command.type === 'paste-internal' || command.type === 'autofill') &&
      command.mode !== 'format'
    ) {
      const source = command.source.range;
      const target = internalPasteRange(
        source,
        command.target.range,
        command.type === 'paste-internal' && command.cut,
      );
      const count =
        (target.end.row - target.start.row + 1) * (target.end.column - target.start.column + 1);
      if (!reserve(count)) {
        return {
          ok: false,
          result: rejected('TRANSACTION_LIMIT_EXCEEDED', 'Validation candidate limit exceeded'),
        };
      }
      for (let row = target.start.row; row <= target.end.row; row += 1) {
        for (let column = target.start.column; column <= target.end.column; column += 1) {
          output.push({ sheet: command.target.sheet, row, column, text: '' });
        }
      }
    }
  }
  return { ok: true, values: output };
}

type CapturedValidatedTransactionRequest = Omit<ValidatedTransactionRequest, 'transaction'> & {
  readonly transaction: DocumentTransactionEnvelope;
};

function captureRequestEnvelope(
  input: ValidatedTransactionRequest,
):
  | { readonly ok: true; readonly request: CapturedValidatedTransactionRequest }
  | { readonly ok: false; readonly result: DocumentTransactionResult } {
  try {
    const descriptors = Object.getOwnPropertyDescriptors(input);
    const data = <Key extends keyof ValidatedTransactionRequest>(
      key: Key,
      required = false,
    ): ValidatedTransactionRequest[Key] => {
      const descriptor = descriptors[key];
      if (descriptor === undefined) {
        if (required) throw new TypeError(`Missing ${String(key)}`);
        return undefined as ValidatedTransactionRequest[Key];
      }
      if (!Object.hasOwn(descriptor, 'value')) {
        throw new TypeError(`${String(key)} must be a data property`);
      }
      return descriptor.value as ValidatedTransactionRequest[Key];
    };
    return {
      ok: true,
      request: Object.freeze({
        controller: data('controller', true),
        transaction: data('transaction', true),
        validation: data('validation'),
        options: data('options'),
        signal: data('signal'),
        confirmWarning: data('confirmWarning'),
        limits: data('limits'),
        canCommit: data('canCommit'),
      }),
    };
  } catch {
    return {
      ok: false,
      result: rejected('COMMAND_SCHEMA_INVALID', 'Validation request could not be isolated'),
    };
  }
}

/** Validates a content batch against one snapshot and commits it as one history entry. */
export async function executeValidatedTransaction(
  input: ValidatedTransactionRequest,
): Promise<DocumentTransactionResult> {
  const envelope = captureRequestEnvelope(input);
  if (!envelope.ok) return envelope.result;
  const request = envelope.request;
  const captured = snapshotSerializableTransaction(
    request.transaction as unknown as SerializableTransactionEnvelope,
  );
  if ('status' in captured) return rejected(captured.code, captured.message);
  const transaction = captured as unknown as DocumentTransactionEnvelope;
  const owner = request.controller;
  const snapshot = owner.getSnapshot();
  if (transaction.baseRevision !== snapshot.revision) {
    return rejected('REVISION_CONFLICT', 'Transaction base revision is stale');
  }
  if (aborted(request.signal)) {
    return rejected('VALIDATION_SOURCE_ABORTED', 'Validation was aborted');
  }
  const maxConcurrency = request.limits?.maxConcurrency ?? 8;
  const totalTimeoutMs = request.limits?.totalTimeoutMs ?? 5_000;
  if (
    !validPositiveLimit(maxConcurrency) ||
    maxConcurrency > 32 ||
    !validPositiveLimit(totalTimeoutMs)
  ) {
    return rejected(
      'VALIDATION_LIMIT_INVALID',
      'Validation limits must be finite positive integers',
    );
  }

  const collected = candidates(transaction);
  if (!collected.ok) return collected.result;
  const preview = previewValidatedDocumentTransaction(owner, transaction, request.options ?? {});
  if (preview.status === 'rejected') {
    return rejected(
      preview.code ?? 'TRANSACTION_INVARIANT_FAILED',
      preview.message ?? 'Invalid transaction',
    );
  }
  const plannedDocument = preview.document ?? snapshot.document;
  const controller = new AbortController();
  const abort = () => controller.abort(request.signal?.reason);
  request.signal?.addEventListener('abort', abort, { once: true });
  const timeout = setTimeout(
    () => controller.abort(new Error('Validation transaction timed out')),
    totalTimeoutMs,
  );
  const engine = createValidationEngine(request.validation);
  try {
    const requests: ValidationRequest[] = [];
    for (const candidate of collected.values) {
      const resolution = resolveDocumentValidation(
        plannedDocument,
        {
          sheetId: candidate.sheet as string as import('../document').DocumentSheetId,
          row: candidate.row,
          column: candidate.column,
        },
        candidateText(plannedDocument, candidate),
        controller.signal,
      );
      if (resolution.kind === 'invalid') {
        return rejected(
          'VALIDATION_RULE_INVALID',
          `Referenced validation rule ${resolution.validationId} is invalid`,
        );
      }
      if (resolution.kind === 'legacy') continue;
      if (resolution.kind === 'request') requests.push(resolution.request);
    }

    const results: ValidationResult[] = [];
    let next = 0;
    const worker = async (): Promise<void> => {
      while (next < requests.length && !controller.signal.aborted) {
        const index = next;
        next += 1;
        results[index] = await engine.validate(requests[index]!);
      }
    };
    try {
      await beforeDeadline(
        Promise.all(
          Array.from({ length: Math.min(maxConcurrency, requests.length) }, () => worker()),
        ).then(() => undefined),
        controller.signal,
      );
    } catch {
      return rejected(
        request.signal?.aborted === true
          ? 'VALIDATION_SOURCE_ABORTED'
          : 'VALIDATION_SOURCE_TIMEOUT',
        'Validation was aborted',
      );
    }
    for (const result of results) {
      if (result.status === 'error') return rejected(result.code, 'Validation could not complete');
      if (result.status === 'rejected') {
        return rejected('VALIDATION_REJECTED', 'Cell value failed validation');
      }
      if (result.status === 'warning') {
        if (request.confirmWarning === undefined) {
          return rejected('VALIDATION_REJECTED', 'Validation warning was not confirmed');
        }
        let confirmed = false;
        try {
          confirmed = await beforeDeadline(
            Promise.resolve(request.confirmWarning(result)),
            controller.signal,
          );
        } catch {
          return rejected(
            request.signal?.aborted === true
              ? 'VALIDATION_SOURCE_ABORTED'
              : 'VALIDATION_SOURCE_TIMEOUT',
            'Validation warning confirmation timed out',
          );
        }
        if (!confirmed) {
          return rejected('VALIDATION_REJECTED', 'Validation warning was not confirmed');
        }
      }
    }

    if (owner.getSnapshot().revision !== snapshot.revision) {
      return rejected('REVISION_CONFLICT', 'Document changed while validation was pending');
    }
    if (request.canCommit?.() === false) {
      return rejected('VALIDATION_EDIT_CANCELLED', 'Validated edit was cancelled before commit');
    }
    return commitValidatedDocumentTransaction(owner, transaction, request.options ?? {});
  } finally {
    engine.dispose?.();
    clearTimeout(timeout);
    request.signal?.removeEventListener('abort', abort);
  }
}
