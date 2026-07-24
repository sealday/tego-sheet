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
  | { readonly status: 'rejected'; readonly operationId: string };

export interface RemoteOperationProcessor {
  readonly revision: string;
  process(operation: RemoteOperation): RemoteOperationOutcome;
}

export interface CreateRemoteOperationProcessorOptions {
  readonly initialRevision: string;
  readonly maximumOperationBytes?: number;
  readonly maximumRememberedOperations?: number;
  readonly apply: (transaction: SerializableTransactionEnvelope) => boolean;
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
  return Object.freeze({
    ...snapshot,
    commands: Object.freeze(snapshot.commands),
    ...(snapshot.metadata === undefined ? {} : { metadata: Object.freeze(snapshot.metadata) }),
  });
}

/** Validates, orders, deduplicates, and atomically applies host-transformed remote transactions. */
export function createRemoteOperationProcessor(
  options: CreateRemoteOperationProcessorOptions,
): RemoteOperationProcessor {
  let revision = identifier(options.initialRevision, 'Collaboration revision');
  const maximumOperationBytes = options.maximumOperationBytes ?? 1024 * 1024;
  const maximumRememberedOperations = options.maximumRememberedOperations ?? 10_000;
  if (!Number.isSafeInteger(maximumOperationBytes) || maximumOperationBytes < 1) {
    throw new RangeError('Collaboration maximumOperationBytes must be a positive safe integer');
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
  return Object.freeze({
    get revision(): string {
      return revision;
    },
    process(operation: RemoteOperation): RemoteOperationOutcome {
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
      if (baseRevision !== revision) {
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
      let accepted = false;
      try {
        accepted = options.apply(transaction);
      } catch {
        accepted = false;
      }
      if (!accepted) return Object.freeze({ status: 'rejected', operationId });
      revision = nextRevision;
      remembered.add(operationId);
      order.push(operationId);
      if (order.length > maximumRememberedOperations) {
        remembered.delete(order.shift() as string);
      }
      return Object.freeze({ status: 'applied', revision });
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
