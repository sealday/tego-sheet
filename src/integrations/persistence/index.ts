import type { Diagnostic } from '../../document';
import type { SerializableTransactionEnvelope } from '../../core/controller/spreadsheet-document-controller';

export type SaveReason = 'manual' | 'autosave' | 'checkpoint' | 'before-close';

/** Immutable persistence request sent to the host adapter. */
export interface SaveRequest {
  readonly documentId: string;
  readonly requestId: string;
  readonly baseRevision: string;
  readonly transactions: readonly SerializableTransactionEnvelope[];
  readonly reason: SaveReason;
}

/** Host persistence acknowledgement or explicit conflict/rejection. */
export type SaveResult =
  | {
      readonly status: 'saved';
      readonly revision: string;
      readonly persistedTransactionIds: readonly string[];
    }
  | { readonly status: 'conflict'; readonly currentRevision: string }
  | { readonly status: 'rejected'; readonly code: string; readonly message: string };

export interface PersistencePort {
  save(request: SaveRequest, signal: AbortSignal): Promise<SaveResult>;
}

export type PersistenceState =
  | { readonly status: 'clean'; readonly revision: string; readonly savedAt: number }
  | {
      readonly status: 'dirty';
      readonly revision: string;
      readonly pending: readonly string[];
    }
  | {
      readonly status: 'saving';
      readonly revision: string;
      readonly requestId: string;
      readonly inFlight: readonly string[];
      readonly pending: readonly string[];
    }
  | {
      readonly status: 'conflict';
      readonly baseRevision: string;
      readonly currentRevision: string;
      readonly pending: readonly string[];
    }
  | {
      readonly status: 'error';
      readonly revision: string;
      readonly pending: readonly string[];
      readonly diagnostic: Diagnostic;
    };

export interface CreatePersistenceControllerOptions {
  readonly documentId: string;
  readonly initialRevision: string;
  readonly adapter: PersistencePort;
  readonly requestId: () => string;
  readonly now?: () => number;
  readonly maximumTransactionsPerSave?: number;
  readonly maximumRequestBytes?: number;
}

export interface PersistenceController {
  readonly state: PersistenceState;
  enqueue(transaction: SerializableTransactionEnvelope): void;
  save(reason?: SaveReason): Promise<SaveResult>;
  retry(): Promise<SaveResult>;
  hasPendingChanges(): boolean;
  dispose(): void;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function deepFreeze<Value>(value: Value): Value {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

function snapshotTransaction(
  transaction: SerializableTransactionEnvelope,
): SerializableTransactionEnvelope {
  let snapshot: SerializableTransactionEnvelope;
  try {
    snapshot = JSON.parse(JSON.stringify(transaction)) as SerializableTransactionEnvelope;
  } catch {
    throw new TypeError('Persistence transaction must be JSON serializable');
  }
  if (
    snapshot.schemaVersion !== 1 ||
    !identifierPattern.test(snapshot.id) ||
    !Number.isSafeInteger(snapshot.baseRevision) ||
    snapshot.baseRevision < 0 ||
    !Array.isArray(snapshot.commands)
  ) {
    throw new TypeError('Persistence transaction is invalid');
  }
  return deepFreeze(snapshot);
}

function frozenIds(transactions: Iterable<SerializableTransactionEnvelope>): readonly string[] {
  return Object.freeze([...transactions].map(({ id }) => id));
}

/** Creates a revision/ack-driven persistence coordinator with isolated in-flight batches. */
export function createPersistenceController(
  options: CreatePersistenceControllerOptions,
): PersistenceController {
  const documentId = identifier(options.documentId, 'Persistence documentId');
  let revision = identifier(options.initialRevision, 'Persistence revision');
  const maximumTransactions = options.maximumTransactionsPerSave ?? 1_000;
  const maximumRequestBytes = options.maximumRequestBytes ?? 8 * 1024 * 1024;
  if (
    !Number.isSafeInteger(maximumTransactions) ||
    maximumTransactions < 1 ||
    maximumTransactions > 1_000
  ) {
    throw new RangeError('Persistence maximumTransactionsPerSave must be from 1 through 1000');
  }
  if (!Number.isSafeInteger(maximumRequestBytes) || maximumRequestBytes < 1) {
    throw new RangeError('Persistence maximumRequestBytes must be a positive safe integer');
  }
  const pending = new Map<string, SerializableTransactionEnvelope>();
  let state: PersistenceState = Object.freeze({
    status: 'clean',
    revision,
    savedAt: options.now?.() ?? Date.now(),
  });
  let disposed = false;
  let activeAbort: AbortController | undefined;
  let activePromise: Promise<SaveResult> | undefined;
  let retryRequest: SaveRequest | undefined;

  const pendingState = (): void => {
    if (activePromise !== undefined && state.status === 'saving') {
      state = Object.freeze({
        ...state,
        pending: frozenIds(
          [...pending.values()].filter(
            ({ id }) => !(state.status === 'saving' && state.inFlight.includes(id)),
          ),
        ),
      });
      return;
    }
    state =
      pending.size === 0
        ? Object.freeze({
            status: 'clean',
            revision,
            savedAt: options.now?.() ?? Date.now(),
          })
        : Object.freeze({
            status: 'dirty',
            revision,
            pending: frozenIds(pending.values()),
          });
  };

  const execute = (request: SaveRequest): Promise<SaveResult> => {
    const inFlightIds = request.transactions.map(({ id }) => id);
    activeAbort = new AbortController();
    state = Object.freeze({
      status: 'saving',
      revision: request.baseRevision,
      requestId: request.requestId,
      inFlight: Object.freeze(inFlightIds),
      pending: frozenIds([...pending.values()].filter(({ id }) => !inFlightIds.includes(id))),
    });
    retryRequest = request;
    const promise = options.adapter
      .save(request, activeAbort.signal)
      .then((result): SaveResult => {
        if (result.status === 'saved') {
          revision = identifier(result.revision, 'Persistence acknowledgement revision');
          const acknowledged = new Set<string>();
          for (const id of result.persistedTransactionIds) {
            if (acknowledged.has(id) || !inFlightIds.includes(id)) {
              throw new TypeError(`Persistence acknowledgement contains invalid transaction ${id}`);
            }
            acknowledged.add(id);
            pending.delete(id);
          }
          retryRequest = undefined;
          return deepFreeze({
            status: 'saved',
            revision,
            persistedTransactionIds: [...acknowledged],
          });
        }
        if (result.status === 'conflict') {
          const currentRevision = identifier(
            result.currentRevision,
            'Persistence conflict revision',
          );
          state = Object.freeze({
            status: 'conflict',
            baseRevision: request.baseRevision,
            currentRevision,
            pending: frozenIds(pending.values()),
          });
          return Object.freeze({ status: 'conflict', currentRevision });
        }
        const diagnostic: Diagnostic = Object.freeze({
          code: result.code || 'PERSISTENCE_SAVE_FAILED',
          severity: 'error',
          domain: 'persistence',
          stage: 'save',
          message: result.message || 'Persistence adapter rejected the save',
        });
        state = Object.freeze({
          status: 'error',
          revision,
          pending: frozenIds(pending.values()),
          diagnostic,
        });
        return Object.freeze({
          status: 'rejected',
          code: diagnostic.code,
          message: diagnostic.message,
        });
      })
      .catch((cause: unknown): never => {
        const diagnostic: Diagnostic = Object.freeze({
          code: activeAbort?.signal.aborted ? 'PERSISTENCE_CANCELLED' : 'PERSISTENCE_SAVE_FAILED',
          severity: 'error',
          domain: 'persistence',
          stage: 'save',
          message: cause instanceof Error ? cause.message : 'Persistence save failed',
          cause,
        });
        state = Object.freeze({
          status: 'error',
          revision,
          pending: frozenIds(pending.values()),
          diagnostic,
        });
        throw cause;
      })
      .finally(() => {
        activePromise = undefined;
        activeAbort = undefined;
        if (state.status === 'saving') pendingState();
      });
    activePromise = promise;
    return promise;
  };

  const controller: PersistenceController = {
    get state(): PersistenceState {
      return state;
    },
    enqueue(transaction): void {
      if (disposed) throw new TypeError('Persistence controller is disposed');
      const snapshot = snapshotTransaction(transaction);
      if (pending.has(snapshot.id)) return;
      pending.set(snapshot.id, snapshot);
      pendingState();
    },
    async save(reason: SaveReason = 'manual'): Promise<SaveResult> {
      if (disposed) throw new TypeError('Persistence controller is disposed');
      if (activePromise !== undefined) return activePromise;
      if (pending.size === 0) {
        return Object.freeze({
          status: 'saved',
          revision,
          persistedTransactionIds: Object.freeze([]),
        });
      }
      const transactions = [...pending.values()].slice(0, maximumTransactions);
      const request: SaveRequest = deepFreeze({
        documentId,
        requestId: identifier(options.requestId(), 'Persistence requestId'),
        baseRevision: revision,
        transactions,
        reason,
      });
      const bytes = new TextEncoder().encode(JSON.stringify(request)).byteLength;
      if (bytes > maximumRequestBytes) {
        throw new RangeError(`Persistence request exceeds ${maximumRequestBytes} bytes`);
      }
      return execute(request);
    },
    async retry(): Promise<SaveResult> {
      if (disposed) throw new TypeError('Persistence controller is disposed');
      if (activePromise !== undefined) return activePromise;
      if (retryRequest === undefined) throw new TypeError('No persistence request is retryable');
      return execute(retryRequest);
    },
    hasPendingChanges(): boolean {
      return pending.size > 0;
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      activeAbort?.abort();
    },
  };
  return Object.freeze(controller);
}
