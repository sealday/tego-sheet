import type { SerializableTransactionEnvelope } from '../../core/controller/spreadsheet-document-controller';

export interface RemoteOperation {
  readonly operationId: string;
  readonly actorId: string;
  readonly baseRevision: string;
  readonly revision: string;
  readonly transaction: SerializableTransactionEnvelope;
}

export type RemoteOperationOutcome =
  | { readonly status: 'applied'; readonly revision: string }
  | { readonly status: 'duplicate'; readonly operationId: string }
  | {
      readonly status: 'resync-required';
      readonly expectedRevision: string;
      readonly receivedBaseRevision: string;
    }
  | { readonly status: 'disconnected'; readonly revision: string }
  | { readonly status: 'rejected'; readonly operationId: string };

export interface RemoteOperationProcessor {
  readonly revision: string;
  process(operation: RemoteOperation): RemoteOperationOutcome;
  disconnect(): void;
  resetAfterResync(revision: string): void;
}

export interface PreparedRemoteTransaction {
  /** Commits all prepared host changes as one indivisible operation. */
  commit(): void;
  /** Restores the exact pre-commit host state; must be safe after a partial commit failure. */
  rollback(): void;
}

export interface RemoteTransactionBoundary {
  /** Performs validation and prepares a reversible commit without mutating host state. */
  prepare(
    transaction: SerializableTransactionEnvelope,
    operation: RemoteOperation,
  ): PreparedRemoteTransaction | undefined;
}

export interface CreateRemoteOperationProcessorOptions {
  readonly initialRevision: string;
  readonly maximumOperationBytes?: number;
  readonly maximumRememberedOperations?: number;
  /** Defaults to deny when omitted. Server authorization remains mandatory. */
  readonly permissionGate?: (
    transaction: SerializableTransactionEnvelope,
    operation: RemoteOperation,
  ) => boolean;
  readonly transactionBoundary: RemoteTransactionBoundary;
}

export interface CollaborationOutboundOperation {
  readonly protocolVersion: 1;
  readonly documentId: string;
  readonly operationId: string;
  readonly baseRevision: string;
  readonly revision: string;
  /** Collaborative undo is transported as an ordinary validated inverse transaction. */
  readonly undoOfOperationId?: string;
  readonly transaction: SerializableTransactionEnvelope;
}

export interface CollaborationCapabilities {
  readonly protocolVersions: readonly number[];
  readonly collaborativeUndo: boolean;
}

export interface CollaborationOutbox {
  put(operation: CollaborationOutboundOperation): Promise<void>;
  remove(operationId: string): Promise<void>;
  list(documentId: string): Promise<readonly CollaborationOutboundOperation[]>;
}

export interface CollaborationOperationPort {
  submit(
    operation: CollaborationOutboundOperation,
    signal: AbortSignal,
  ): Promise<{ readonly operationId: string }>;
}

export interface CollaborationOutboxCoordinator {
  queue(
    operation: CollaborationOutboundOperation,
    signal: AbortSignal,
  ): Promise<
    | { readonly status: 'queued'; readonly operationId: string }
    | { readonly status: 'submitted'; readonly operationId: string }
  >;
  disconnect(): void;
  reconnect(
    capabilities: CollaborationCapabilities,
    signal: AbortSignal,
  ): Promise<readonly string[]>;
}

export interface RemotePresence {
  readonly actorId: string;
  readonly sheetId: string;
  readonly activeCell?: { readonly row: number; readonly column: number };
  readonly selections: readonly {
    readonly start: { readonly row: number; readonly column: number };
    readonly end: { readonly row: number; readonly column: number };
  }[];
  readonly display: { readonly label: string; readonly color: string };
  readonly expiresAt: number;
}

export interface PresenceStore {
  replace(presence: readonly RemotePresence[]): void;
  list(): readonly RemotePresence[];
  close(): void;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function snapshotTransaction(
  transaction: SerializableTransactionEnvelope,
): SerializableTransactionEnvelope {
  const snapshot = JSON.parse(JSON.stringify(transaction)) as SerializableTransactionEnvelope;
  if (
    snapshot.schemaVersion !== 1 ||
    !identifierPattern.test(snapshot.id) ||
    !Number.isSafeInteger(snapshot.baseRevision) ||
    snapshot.baseRevision < 0 ||
    !Array.isArray(snapshot.commands)
  ) {
    throw new TypeError('Remote transaction is invalid');
  }
  if (snapshot.commands.some(({ command }) => command.type === 'undo' || command.type === 'redo')) {
    throw new TypeError(
      'Remote transaction cannot invoke local history; send an inverse transaction for undo',
    );
  }
  return deepFreeze(snapshot);
}

function deepFreeze<Value>(value: Value): Value {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

/** Validates, orders, deduplicates, and atomically applies host-transformed remote transactions. */
export function createRemoteOperationProcessor(
  options: CreateRemoteOperationProcessorOptions,
): RemoteOperationProcessor {
  let revision = identifier(options.initialRevision, 'Collaboration revision');
  const maximumOperationBytes = options.maximumOperationBytes ?? 1024 * 1024;
  const maximumRememberedOperations = options.maximumRememberedOperations ?? 10_000;
  if (
    !Number.isSafeInteger(maximumOperationBytes) ||
    maximumOperationBytes < 1 ||
    maximumOperationBytes > 1024 * 1024
  ) {
    throw new RangeError('Collaboration maximumOperationBytes must be from 1 byte through 1 MiB');
  }
  if (
    !Number.isSafeInteger(maximumRememberedOperations) ||
    maximumRememberedOperations < 1 ||
    maximumRememberedOperations > 10_000
  ) {
    throw new RangeError('Collaboration maximumRememberedOperations must be from 1 through 10000');
  }
  const remembered = new Set<string>();
  const order: string[] = [];
  let connected = true;
  let resyncRequired = false;
  return Object.freeze({
    get revision(): string {
      return revision;
    },
    process(operation: RemoteOperation): RemoteOperationOutcome {
      if (!connected) return Object.freeze({ status: 'disconnected', revision });
      const operationId = identifier(operation.operationId, 'Collaboration operationId');
      if (remembered.has(operationId)) {
        return Object.freeze({ status: 'duplicate', operationId });
      }
      const baseRevision = identifier(
        operation.baseRevision,
        'Collaboration operation baseRevision',
      );
      const nextRevision = identifier(operation.revision, 'Collaboration operation revision');
      identifier(operation.actorId, 'Collaboration actorId');
      if (resyncRequired || baseRevision !== revision) {
        resyncRequired = true;
        return Object.freeze({
          status: 'resync-required',
          expectedRevision: revision,
          receivedBaseRevision: baseRevision,
        });
      }
      if (nextRevision === baseRevision) {
        return Object.freeze({ status: 'rejected', operationId });
      }
      const bytes = new TextEncoder().encode(JSON.stringify(operation)).byteLength;
      if (bytes > maximumOperationBytes) {
        throw new RangeError(
          `Collaboration operation cannot exceed ${maximumOperationBytes} bytes`,
        );
      }
      const transaction = snapshotTransaction(operation.transaction);
      if (!(options.permissionGate?.(transaction, operation) ?? false)) {
        return Object.freeze({ status: 'rejected', operationId });
      }
      const prepared = options.transactionBoundary.prepare(transaction, operation);
      if (prepared === undefined) return Object.freeze({ status: 'rejected', operationId });
      try {
        prepared.commit();
      } catch {
        try {
          prepared.rollback();
        } catch {
          // The host boundary owns rollback durability. Never advance protocol state on failure.
        }
        return Object.freeze({ status: 'rejected', operationId });
      }
      revision = nextRevision;
      remembered.add(operationId);
      order.push(operationId);
      if (order.length > maximumRememberedOperations) {
        remembered.delete(order.shift() as string);
      }
      return Object.freeze({ status: 'applied', revision });
    },
    disconnect(): void {
      connected = false;
    },
    resetAfterResync(nextRevision: string): void {
      revision = identifier(nextRevision, 'Collaboration resync revision');
      remembered.clear();
      order.length = 0;
      resyncRequired = false;
      connected = true;
    },
  });
}

function snapshotOutboundOperation(
  operation: CollaborationOutboundOperation,
): CollaborationOutboundOperation {
  let snapshot: CollaborationOutboundOperation;
  try {
    snapshot = JSON.parse(JSON.stringify(operation)) as CollaborationOutboundOperation;
  } catch {
    throw new TypeError('Collaboration outbound operation must be JSON serializable');
  }
  if (snapshot.protocolVersion !== 1) {
    throw new TypeError('Collaboration outbound protocol version is unsupported');
  }
  if (new TextEncoder().encode(JSON.stringify(snapshot)).byteLength > 1024 * 1024) {
    throw new RangeError('Collaboration outbound operation cannot exceed 1 MiB');
  }
  for (const [label, value] of [
    ['documentId', snapshot.documentId],
    ['operationId', snapshot.operationId],
    ['baseRevision', snapshot.baseRevision],
    ['revision', snapshot.revision],
  ] as const) {
    identifier(value, `Collaboration outbound ${label}`);
  }
  if (snapshot.baseRevision === snapshot.revision) {
    throw new TypeError('Collaboration outbound revision must advance');
  }
  if (snapshot.undoOfOperationId !== undefined) {
    identifier(snapshot.undoOfOperationId, 'Collaboration undo operationId');
    if (snapshot.undoOfOperationId === snapshot.operationId) {
      throw new TypeError('Collaboration operation cannot undo itself');
    }
  }
  const transaction = snapshotTransaction(snapshot.transaction);
  return deepFreeze({ ...snapshot, transaction });
}

/** Coordinates a host-owned durable local-operation outbox and explicit replay negotiation. */
export function createCollaborationOutboxCoordinator(options: {
  readonly documentId: string;
  readonly outbox: CollaborationOutbox;
  readonly adapter: CollaborationOperationPort;
}): CollaborationOutboxCoordinator {
  const documentId = identifier(options.documentId, 'Collaboration outbox documentId');
  let connected = true;
  let capabilities: CollaborationCapabilities | undefined = {
    protocolVersions: Object.freeze([1]),
    collaborativeUndo: true,
  };
  const activeSubmissions = new Set<AbortController>();

  const assertCapabilities = (
    operation: CollaborationOutboundOperation,
    current: CollaborationCapabilities,
  ): void => {
    if (!current.protocolVersions.includes(operation.protocolVersion)) {
      throw new TypeError('Collaboration replay protocol capability is unavailable');
    }
    if (operation.undoOfOperationId !== undefined && !current.collaborativeUndo) {
      throw new TypeError('Collaboration replay undo capability is unavailable');
    }
  };

  const submitStored = async (
    operation: CollaborationOutboundOperation,
    signal: AbortSignal,
    current: CollaborationCapabilities,
  ): Promise<string> => {
    if (signal.aborted) throw new TypeError('Collaboration operation was cancelled');
    assertCapabilities(operation, current);
    const controller = new AbortController();
    const abort = (): void => controller.abort();
    signal.addEventListener('abort', abort, { once: true });
    activeSubmissions.add(controller);
    try {
      const acknowledgement = await options.adapter.submit(operation, controller.signal);
      if (acknowledgement.operationId !== operation.operationId) {
        throw new TypeError('Collaboration acknowledgement does not match operationId');
      }
      await options.outbox.remove(operation.operationId);
      return acknowledgement.operationId;
    } finally {
      activeSubmissions.delete(controller);
      signal.removeEventListener('abort', abort);
    }
  };

  return Object.freeze({
    async queue(
      operation: CollaborationOutboundOperation,
      signal: AbortSignal,
    ): Promise<
      | { readonly status: 'queued'; readonly operationId: string }
      | { readonly status: 'submitted'; readonly operationId: string }
    > {
      const snapshot = snapshotOutboundOperation(operation);
      if (snapshot.documentId !== documentId) {
        throw new TypeError('Collaboration outbound documentId does not match coordinator');
      }
      if (signal.aborted) throw new TypeError('Collaboration operation was cancelled');
      await options.outbox.put(snapshot);
      if (!connected || capabilities === undefined) {
        return Object.freeze({ status: 'queued', operationId: snapshot.operationId });
      }
      await submitStored(snapshot, signal, capabilities);
      return Object.freeze({ status: 'submitted', operationId: snapshot.operationId });
    },
    disconnect(): void {
      connected = false;
      capabilities = undefined;
      for (const controller of activeSubmissions) controller.abort();
      activeSubmissions.clear();
    },
    async reconnect(
      nextCapabilities: CollaborationCapabilities,
      signal: AbortSignal,
    ): Promise<readonly string[]> {
      const normalizedCapabilities: CollaborationCapabilities = deepFreeze({
        protocolVersions: [...nextCapabilities.protocolVersions],
        collaborativeUndo: nextCapabilities.collaborativeUndo,
      });
      if (
        normalizedCapabilities.protocolVersions.length === 0 ||
        normalizedCapabilities.protocolVersions.some(
          (version) => !Number.isSafeInteger(version) || version < 1,
        )
      ) {
        throw new TypeError('Collaboration replay protocol capabilities are invalid');
      }
      const stored = (await options.outbox.list(documentId)).map(snapshotOutboundOperation);
      const seen = new Set<string>();
      for (const [index, operation] of stored.entries()) {
        if (operation.documentId !== documentId || seen.has(operation.operationId)) {
          throw new TypeError('Collaboration replay identity is invalid');
        }
        const previous = stored[index - 1];
        if (previous !== undefined && previous.revision !== operation.baseRevision) {
          throw new TypeError('Collaboration replay revision chain is not contiguous');
        }
        assertCapabilities(operation, normalizedCapabilities);
        seen.add(operation.operationId);
      }
      capabilities = normalizedCapabilities;
      connected = true;
      const acknowledged: string[] = [];
      for (const operation of stored) {
        acknowledged.push(await submitStored(operation, signal, normalizedCapabilities));
      }
      return Object.freeze(acknowledged);
    },
  });
}

function coordinate(value: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new TypeError(`${label} must be a non-negative safe integer`);
  }
  return value;
}

function snapshotPresence(presence: RemotePresence): RemotePresence {
  const label = [...presence.display.label]
    .filter((character) => {
      const code = character.codePointAt(0) ?? 0;
      return code > 31 && code !== 127;
    })
    .join('')
    .slice(0, 256);
  if (!/^#[0-9a-f]{6}$/iu.test(presence.display.color)) {
    throw new TypeError('Collaboration presence color is invalid');
  }
  if (!Number.isFinite(presence.expiresAt) || presence.expiresAt < 0) {
    throw new TypeError('Collaboration presence expiry is invalid');
  }
  return Object.freeze({
    actorId: identifier(presence.actorId, 'Collaboration presence actorId'),
    sheetId: identifier(presence.sheetId, 'Collaboration presence sheetId'),
    ...(presence.activeCell === undefined
      ? {}
      : {
          activeCell: Object.freeze({
            row: coordinate(presence.activeCell.row, 'Presence active row'),
            column: coordinate(presence.activeCell.column, 'Presence active column'),
          }),
        }),
    selections: Object.freeze(
      presence.selections.map((selection) =>
        Object.freeze({
          start: Object.freeze({
            row: coordinate(selection.start.row, 'Presence selection row'),
            column: coordinate(selection.start.column, 'Presence selection column'),
          }),
          end: Object.freeze({
            row: coordinate(selection.end.row, 'Presence selection row'),
            column: coordinate(selection.end.column, 'Presence selection column'),
          }),
        }),
      ),
    ),
    display: Object.freeze({ label, color: presence.display.color.toLowerCase() }),
    expiresAt: presence.expiresAt,
  });
}

/** Creates an ephemeral presence store that never references the workbook document. */
export function createPresenceStore(options: { readonly now?: () => number } = {}): PresenceStore {
  let closed = false;
  let current: readonly RemotePresence[] = Object.freeze([]);
  return Object.freeze({
    replace(presence: readonly RemotePresence[]): void {
      if (closed) throw new TypeError('Presence store is closed');
      if (presence.length > 10_000) throw new RangeError('Presence participant limit is 10000');
      const actorIds = new Set<string>();
      current = Object.freeze(
        presence.map((entry) => {
          const snapshot = snapshotPresence(entry);
          if (actorIds.has(snapshot.actorId)) {
            throw new TypeError(`Duplicate presence actor ${snapshot.actorId}`);
          }
          actorIds.add(snapshot.actorId);
          return snapshot;
        }),
      );
    },
    list(): readonly RemotePresence[] {
      if (closed) return Object.freeze([]);
      const now = options.now?.() ?? Date.now();
      current = Object.freeze(current.filter(({ expiresAt }) => expiresAt > now));
      return current;
    },
    close(): void {
      closed = true;
      current = Object.freeze([]);
    },
  });
}
