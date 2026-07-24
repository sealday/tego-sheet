import type { CoordinateTransform } from '../../core/coordinates/coordinate-transform';

export interface CommentCellAddress {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

export interface CommentRange {
  readonly sheetId: string;
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
}

export type CommentAnchorLocation =
  | { readonly type: 'cell'; readonly cell: CommentCellAddress }
  | { readonly type: 'range'; readonly range: CommentRange }
  | { readonly type: 'object'; readonly sheetId: string; readonly objectId: string };

export type CommentAnchor =
  | CommentAnchorLocation
  | {
      readonly type: 'orphaned';
      readonly lastKnown: CommentAnchorLocation;
      readonly reason: string;
    };

export interface CommentAnchorUpdate {
  readonly threadId: string;
  readonly expectedThreadRevision: string;
  readonly anchor: CommentAnchor;
}

export interface CommentAnchorUpdateBatch {
  readonly operationId: string;
  readonly documentId: string;
  readonly fromDocumentRevision: string;
  readonly toDocumentRevision: string;
  readonly updates: readonly CommentAnchorUpdate[];
}

export interface CommentAnchorUpdateAck {
  readonly operationId: string;
}

/** Host-owned durable storage. `put` must complete before a remote submission begins. */
export interface CommentAnchorOutbox {
  put(batch: CommentAnchorUpdateBatch): Promise<void>;
  remove(operationId: string): Promise<void>;
  list(documentId: string): Promise<readonly CommentAnchorUpdateBatch[]>;
}

export interface CommentAnchorUpdatePort {
  submit(batch: CommentAnchorUpdateBatch, signal: AbortSignal): Promise<CommentAnchorUpdateAck>;
}

export interface CommentAnchorOutboxCoordinator {
  queue(batch: CommentAnchorUpdateBatch, signal: AbortSignal): Promise<CommentAnchorUpdateAck>;
  resume(documentId: string, signal: AbortSignal): Promise<readonly CommentAnchorUpdateAck[]>;
}

export interface CommentAnchorTransformRequest {
  readonly sheetId: string;
  readonly transform: CoordinateTransform;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;

function deepFreeze<Value>(value: Value): Value {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

function snapshotBatch(batch: CommentAnchorUpdateBatch): CommentAnchorUpdateBatch {
  let snapshot: CommentAnchorUpdateBatch;
  try {
    snapshot = JSON.parse(JSON.stringify(batch)) as CommentAnchorUpdateBatch;
  } catch {
    throw new TypeError('Comment anchor update batch must be JSON serializable');
  }
  for (const [label, value] of [
    ['operationId', snapshot.operationId],
    ['documentId', snapshot.documentId],
    ['fromDocumentRevision', snapshot.fromDocumentRevision],
    ['toDocumentRevision', snapshot.toDocumentRevision],
  ] as const) {
    if (!identifierPattern.test(value)) throw new TypeError(`Comment batch ${label} is invalid`);
  }
  if (!Array.isArray(snapshot.updates) || snapshot.updates.length > 10_000) {
    throw new RangeError('Comment anchor update limit is 10000');
  }
  const threads = new Set<string>();
  for (const update of snapshot.updates) {
    if (
      !identifierPattern.test(update.threadId) ||
      !identifierPattern.test(update.expectedThreadRevision) ||
      threads.has(update.threadId)
    ) {
      throw new TypeError('Comment anchor update thread identity is invalid or duplicated');
    }
    threads.add(update.threadId);
  }
  return deepFreeze(snapshot);
}

/** Applies one row/column transform, preserving unrelated sheets and orphaning deleted targets. */
export function transformCommentAnchor(
  anchor: CommentAnchor,
  request: CommentAnchorTransformRequest,
): CommentAnchor {
  if (anchor.type === 'orphaned' || anchor.type === 'object') return anchor;
  if (anchor.type === 'cell') {
    if (anchor.cell.sheetId !== request.sheetId) return anchor;
    const next = request.transform.point(anchor.cell);
    if (next !== null) {
      return deepFreeze({
        type: 'cell',
        cell: { sheetId: anchor.cell.sheetId, ...next },
      });
    }
  } else {
    if (anchor.range.sheetId !== request.sheetId) return anchor;
    const next = request.transform.range(anchor.range);
    if (next !== null) {
      return deepFreeze({
        type: 'range',
        range: { sheetId: anchor.range.sheetId, ...next },
      });
    }
  }
  return deepFreeze({
    type: 'orphaned',
    lastKnown: anchor,
    reason: 'anchor target was deleted by a structural transaction',
  });
}

/** Coordinates a durable idempotent outbox around the host comment adapter. */
export function createCommentAnchorOutboxCoordinator(options: {
  readonly outbox: CommentAnchorOutbox;
  readonly adapter: CommentAnchorUpdatePort;
}): CommentAnchorOutboxCoordinator {
  const submitStored = async (
    batch: CommentAnchorUpdateBatch,
    signal: AbortSignal,
  ): Promise<CommentAnchorUpdateAck> => {
    if (signal.aborted) throw new TypeError('Comment anchor update was cancelled');
    const acknowledgement = await options.adapter.submit(batch, signal);
    if (acknowledgement.operationId !== batch.operationId) {
      throw new TypeError('Comment anchor update acknowledgement does not match operationId');
    }
    await options.outbox.remove(batch.operationId);
    return Object.freeze({ operationId: acknowledgement.operationId });
  };
  return Object.freeze({
    async queue(
      batch: CommentAnchorUpdateBatch,
      signal: AbortSignal,
    ): Promise<CommentAnchorUpdateAck> {
      const snapshot = snapshotBatch(batch);
      if (signal.aborted) throw new TypeError('Comment anchor update was cancelled');
      await options.outbox.put(snapshot);
      return submitStored(snapshot, signal);
    },
    async resume(
      documentId: string,
      signal: AbortSignal,
    ): Promise<readonly CommentAnchorUpdateAck[]> {
      if (!identifierPattern.test(documentId)) throw new TypeError('Comment documentId is invalid');
      const stored = await options.outbox.list(documentId);
      if (stored.length > 10_000) throw new RangeError('Comment outbox resume limit is 10000');
      const acknowledgements: CommentAnchorUpdateAck[] = [];
      for (const batch of stored) {
        acknowledgements.push(await submitStored(snapshotBatch(batch), signal));
      }
      return Object.freeze(acknowledgements);
    },
  });
}
