import type { Diagnostic } from '../../document';
import { parseSpreadsheetDocument, type SpreadsheetDocument } from '../../document';
import type { SerializableTransactionEnvelope } from '../../core/controller/spreadsheet-document-controller';
import type { SpreadsheetDocumentController } from '../../core/controller/spreadsheet-document-controller';

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

export interface PersistedDocumentEnvelope {
  readonly schemaVersion: 1;
  readonly documentId: string;
  readonly revision: string;
  readonly document: SpreadsheetDocument;
}

export interface PersistenceLoadPort {
  load(documentId: string, signal: AbortSignal): Promise<unknown>;
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
  resolveConflict(revision: string): void;
  hasPendingChanges(): boolean;
  dispose(): void;
}

export type PersistenceSessionState =
  | PersistenceState
  | {
      readonly status: 'offline';
      readonly pending: readonly string[];
      readonly revision: string;
    };

export interface PersistenceSession {
  readonly state: PersistenceSessionState;
  attachController(controller: SpreadsheetDocumentController): () => void;
  save(reason?: SaveReason): Promise<SaveResult | { readonly status: 'offline' }>;
  retry(): Promise<SaveResult | { readonly status: 'offline' }>;
  setOnline(online: boolean): void;
  resolveConflict(revision: string): void;
  bindBeforeUnload(target: {
    addEventListener(type: 'beforeunload', listener: (event: BeforeUnloadEvent) => void): void;
    removeEventListener(type: 'beforeunload', listener: (event: BeforeUnloadEvent) => void): void;
  }): () => void;
  subscribe(listener: () => void): () => void;
  dispose(): void;
}

export interface CreatePersistenceSessionOptions extends CreatePersistenceControllerOptions {
  readonly autosaveDelayMs?: number;
  readonly autosaveMaxWaitMs?: number;
  readonly setTimer?: (callback: () => void, delay: number) => unknown;
  readonly clearTimer?: (handle: unknown) => void;
  readonly initiallyOnline?: boolean;
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

/** Loads one host envelope and validates the complete document before exposing it. */
export async function loadPersistedDocument(
  adapter: PersistenceLoadPort,
  documentId: string,
  signal: AbortSignal,
): Promise<PersistedDocumentEnvelope> {
  const expectedDocumentId = identifier(documentId, 'Persistence documentId');
  if (signal.aborted) throw new TypeError('Persistence load was cancelled');
  const value = await adapter.load(expectedDocumentId, signal);
  if (signal.aborted) throw new TypeError('Persistence load was cancelled');
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Persistence load envelope is invalid');
  }
  let snapshot: {
    readonly schemaVersion?: unknown;
    readonly documentId?: unknown;
    readonly revision?: unknown;
    readonly document?: unknown;
  };
  try {
    snapshot = JSON.parse(JSON.stringify(value)) as typeof snapshot;
  } catch {
    throw new TypeError('Persistence load envelope must be JSON serializable');
  }
  if (
    snapshot.schemaVersion !== 1 ||
    snapshot.documentId !== expectedDocumentId ||
    typeof snapshot.revision !== 'string'
  ) {
    throw new TypeError('Persistence load envelope identity is invalid');
  }
  const revision = identifier(snapshot.revision, 'Persistence loaded revision');
  const parsed = parseSpreadsheetDocument(snapshot.document);
  if (!parsed.ok) throw new TypeError('Persistence loaded document is invalid');
  if (parsed.document.id !== expectedDocumentId) {
    throw new TypeError('Persistence loaded documentId does not match envelope');
  }
  return deepFreeze({
    schemaVersion: 1,
    documentId: expectedDocumentId,
    revision,
    document: parsed.document,
  });
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
    if (state.status === 'conflict') return;
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
          const acknowledged = new Set<string>();
          for (const id of result.persistedTransactionIds) {
            if (acknowledged.has(id) || !inFlightIds.includes(id)) {
              throw new TypeError(`Persistence acknowledgement contains invalid transaction ${id}`);
            }
            acknowledged.add(id);
          }
          if (
            acknowledged.size !== inFlightIds.length ||
            inFlightIds.some((id) => !acknowledged.has(id))
          ) {
            throw new TypeError(
              'Persistence acknowledgement must include the complete in-flight batch',
            );
          }
          const acknowledgedRevision = identifier(
            result.revision,
            'Persistence acknowledgement revision',
          );
          for (const id of acknowledged) pending.delete(id);
          revision = acknowledgedRevision;
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
          retryRequest = undefined;
          return Object.freeze({ status: 'conflict', currentRevision });
        }
        const diagnostic: Diagnostic = Object.freeze({
          code: 'PERSISTENCE_SAVE_REJECTED',
          severity: 'error',
          domain: 'persistence',
          stage: 'save',
          message: 'Persistence adapter rejected the save',
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
        const cancelled = activeAbort?.signal.aborted ?? false;
        const diagnostic: Diagnostic = Object.freeze({
          code: cancelled ? 'PERSISTENCE_CANCELLED' : 'PERSISTENCE_SAVE_FAILED',
          severity: 'error',
          domain: 'persistence',
          stage: 'save',
          message: cancelled ? 'Persistence save was cancelled' : 'Persistence save failed',
        });
        state = Object.freeze({
          status: 'error',
          revision,
          pending: frozenIds(pending.values()),
          diagnostic,
        });
        if (cause instanceof TypeError && cause.message.startsWith('Persistence acknowledgement')) {
          throw cause;
        }
        throw new TypeError(diagnostic.message);
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
      if (state.status === 'conflict') {
        throw new TypeError('Persistence conflict must be explicitly resolved before saving');
      }
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
    resolveConflict(nextRevision: string): void {
      if (disposed) throw new TypeError('Persistence controller is disposed');
      if (activePromise !== undefined) {
        throw new TypeError('Cannot resolve a persistence conflict during an active save');
      }
      if (state.status !== 'conflict') {
        throw new TypeError('No persistence conflict is available to resolve');
      }
      revision = identifier(nextRevision, 'Persistence resolved revision');
      retryRequest = undefined;
      state = Object.freeze({
        status: 'dirty',
        revision,
        pending: frozenIds(pending.values()),
      });
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

/** Binds revision persistence to controller commits without owning storage or network policy. */
export function createPersistenceSession(
  options: CreatePersistenceSessionOptions,
): PersistenceSession {
  const persistence = createPersistenceController(options);
  const autosaveDelayMs = options.autosaveDelayMs ?? 1_000;
  const autosaveMaxWaitMs = options.autosaveMaxWaitMs ?? Math.max(10_000, autosaveDelayMs);
  if (!Number.isSafeInteger(autosaveDelayMs) || autosaveDelayMs < 250 || autosaveDelayMs > 60_000) {
    throw new RangeError('Persistence autosaveDelayMs must be from 250 through 60000');
  }
  if (
    !Number.isSafeInteger(autosaveMaxWaitMs) ||
    autosaveMaxWaitMs < autosaveDelayMs ||
    autosaveMaxWaitMs > 60_000
  ) {
    throw new RangeError('Persistence autosaveMaxWaitMs must be from delay through 60000');
  }
  const setTimer = options.setTimer ?? ((callback, delay) => setTimeout(callback, delay));
  const clearTimer =
    options.clearTimer ?? ((handle) => clearTimeout(handle as ReturnType<typeof setTimeout>));
  let online = options.initiallyOnline ?? true;
  let disposed = false;
  let timer: unknown;
  let firstPendingAt: number | undefined;
  let detachController: (() => void) | undefined;
  const listeners = new Set<() => void>();
  const beforeUnloadBindings = new Set<() => void>();
  const now = options.now ?? Date.now;

  const computeState = (): PersistenceSessionState => {
    if (online) return persistence.state;
    const pending =
      persistence.state.status === 'clean' ? Object.freeze([]) : persistence.state.pending;
    return Object.freeze({
      status: 'offline',
      pending,
      revision:
        persistence.state.status === 'conflict'
          ? persistence.state.baseRevision
          : persistence.state.revision,
    });
  };
  let currentState = computeState();
  const publish = (): void => {
    currentState = computeState();
    for (const listener of listeners) listener();
  };
  const clearAutosave = (): void => {
    if (timer === undefined) return;
    clearTimer(timer);
    timer = undefined;
  };
  const scheduleAutosave = (): void => {
    clearAutosave();
    if (persistence.state.status === 'conflict' || persistence.state.status === 'error') {
      return;
    }
    if (!online || !persistence.hasPendingChanges()) {
      if (!persistence.hasPendingChanges()) firstPendingAt = undefined;
      return;
    }
    firstPendingAt ??= now();
    const remaining = Math.max(0, autosaveMaxWaitMs - (now() - firstPendingAt));
    const delay = Math.min(autosaveDelayMs, remaining);
    timer = setTimer(() => {
      timer = undefined;
      const saving = persistence.save('autosave');
      publish();
      void saving.then(
        () => {
          publish();
          if (persistence.state.status === 'dirty') {
            firstPendingAt = now();
            scheduleAutosave();
          } else if (persistence.state.status === 'clean') {
            firstPendingAt = undefined;
          }
        },
        () => {
          publish();
        },
      );
    }, delay);
  };

  const session: PersistenceSession = {
    get state(): PersistenceSessionState {
      return currentState;
    },
    attachController(controller): () => void {
      if (disposed) throw new TypeError('Persistence session is disposed');
      detachController?.();
      const unsubscribe = controller.subscribe(({ snapshot, commit }) => {
        const transaction =
          commit.transaction ??
          ({
            schemaVersion: 1,
            id: `ui:${commit.change.id}`,
            baseRevision: Math.max(0, snapshot.revision - 1),
            commands: [
              {
                schemaVersion: 1,
                id: `command:${commit.change.id}`,
                command: commit.command,
              },
            ],
          } satisfies SerializableTransactionEnvelope);
        persistence.enqueue(transaction);
        scheduleAutosave();
        publish();
      });
      let active = true;
      const detach = (): void => {
        if (!active) return;
        active = false;
        unsubscribe();
        if (detachController === detach) detachController = undefined;
      };
      detachController = detach;
      return detach;
    },
    async save(reason = 'manual') {
      if (disposed) throw new TypeError('Persistence session is disposed');
      clearAutosave();
      if (!online) return Object.freeze({ status: 'offline' as const });
      const saving = persistence.save(reason);
      publish();
      try {
        return await saving;
      } finally {
        publish();
        scheduleAutosave();
      }
    },
    async retry() {
      if (disposed) throw new TypeError('Persistence session is disposed');
      clearAutosave();
      if (!online) return Object.freeze({ status: 'offline' as const });
      const retrying = persistence.retry();
      publish();
      try {
        return await retrying;
      } finally {
        publish();
        scheduleAutosave();
      }
    },
    setOnline(nextOnline): void {
      if (disposed) throw new TypeError('Persistence session is disposed');
      if (online === nextOnline) return;
      online = nextOnline;
      if (online) scheduleAutosave();
      else clearAutosave();
      publish();
    },
    resolveConflict(revision): void {
      persistence.resolveConflict(revision);
      scheduleAutosave();
      publish();
    },
    bindBeforeUnload(target): () => void {
      if (disposed) throw new TypeError('Persistence session is disposed');
      const listener = (event: BeforeUnloadEvent): void => {
        if (!persistence.hasPendingChanges()) return;
        event.preventDefault();
        event.returnValue = '';
      };
      target.addEventListener('beforeunload', listener);
      let active = true;
      const unbind = (): void => {
        if (!active) return;
        active = false;
        target.removeEventListener('beforeunload', listener);
        beforeUnloadBindings.delete(unbind);
      };
      beforeUnloadBindings.add(unbind);
      return unbind;
    },
    subscribe(listener): () => void {
      if (disposed) return () => undefined;
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      clearAutosave();
      detachController?.();
      for (const unbind of beforeUnloadBindings) unbind();
      listeners.clear();
      persistence.dispose();
    },
  };
  return Object.freeze(session);
}
