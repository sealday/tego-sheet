import type { CoordinateTransform } from '../../core/coordinates/coordinate-transform';
import type { PermissionSnapshot } from '../permission';

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

export interface CommentAnchorUpdateConflict {
  readonly status: 'conflict';
  readonly operationId: string;
  readonly expectedRevision: string;
  readonly currentRevision: string;
}

export interface CommentAnchorConflictRecovery {
  readonly action: 'rebase-and-retry';
  readonly rebase: {
    readonly fromRevision: string;
    readonly toRevision: string;
  };
  readonly retry: {
    readonly method: 'queue';
    readonly operationId: string;
  };
}

/** Stable typed failure returned when a durable comment batch must be rebased before retrying. */
export class CommentAnchorOutboxConflictError extends Error {
  readonly code = 'COMMENT_ANCHOR_OUTBOX_CONFLICT';
  readonly operationId: string;
  readonly expectedRevision: string;
  readonly currentRevision: string;
  readonly recovery: CommentAnchorConflictRecovery;

  constructor(
    operationId: string,
    expectedRevision: string,
    currentRevision: string,
    cause?: unknown,
  ) {
    super(
      `Comment anchor update ${operationId} conflicted: expected ${expectedRevision}, current ${currentRevision}; rebase and retry`,
      cause === undefined ? undefined : { cause },
    );
    this.name = 'CommentAnchorOutboxConflictError';
    this.operationId = operationId;
    this.expectedRevision = expectedRevision;
    this.currentRevision = currentRevision;
    this.recovery = deepFreeze({
      action: 'rebase-and-retry',
      rebase: { fromRevision: expectedRevision, toRevision: currentRevision },
      retry: { method: 'queue', operationId },
    });
    Object.freeze(this);
  }
}

/** Host-owned durable storage. `put` must complete before a remote submission begins. */
export interface CommentAnchorOutbox {
  put(batch: CommentAnchorUpdateBatch): Promise<void>;
  remove(operationId: string): Promise<void>;
  list(documentId: string): Promise<readonly CommentAnchorUpdateBatch[]>;
}

export interface CommentAnchorUpdatePort {
  submit(
    batch: CommentAnchorUpdateBatch,
    signal: AbortSignal,
  ): Promise<CommentAnchorUpdateAck | CommentAnchorUpdateConflict>;
}

export interface CommentAnchorOutboxCoordinator {
  queue(batch: CommentAnchorUpdateBatch, signal: AbortSignal): Promise<CommentAnchorUpdateAck>;
  resume(documentId: string, signal: AbortSignal): Promise<readonly CommentAnchorUpdateAck[]>;
}

export interface CommentAnchorTransformRequest {
  readonly sheetId: string;
  readonly transform: CoordinateTransform;
}

export interface CommentRichTextNode {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly code?: boolean;
}

export interface CommentMessage {
  readonly id: string;
  readonly authorId: string;
  readonly content: readonly CommentRichTextNode[];
}

export interface CommentThread {
  readonly id: string;
  readonly revision: string;
  readonly anchor: CommentAnchor;
  readonly messages: readonly CommentMessage[];
  readonly resolved: boolean;
}

export interface CommentStore {
  list(): readonly CommentThread[];
  getSnapshot(): readonly CommentThread[];
  subscribe(listener: () => void): () => void;
  /** Re-evaluates visibility against the latest permission snapshot. */
  refresh(): void;
  create(input: {
    readonly anchor: CommentAnchor;
    readonly content: readonly CommentRichTextNode[];
    readonly expectedDocumentRevision: string;
    readonly currentDocumentRevision: string;
  }): CommentThread;
  reply(
    threadId: string,
    expectedThreadRevision: string,
    content: readonly CommentRichTextNode[],
  ): CommentThread;
  resolve(threadId: string, expectedThreadRevision: string, resolved: boolean): CommentThread;
  remove(threadId: string, expectedThreadRevision: string): void;
  rebase(request: CommentAnchorTransformRequest): readonly CommentThread[];
}

export type CommentPrintPolicy = 'exclude' | 'markers' | 'full';

export interface CommentPrintContent {
  readonly threadId: string;
  readonly marker: string;
  readonly anchor: CommentAnchor;
  readonly text?: string;
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

/** Reduces hostile rich text to the supported text/mark subset. */
export function sanitizeCommentRichText(
  input: readonly CommentRichTextNode[],
): readonly CommentRichTextNode[] {
  if (!Array.isArray(input) || input.length === 0 || input.length > 1_000) {
    throw new RangeError('Comment rich-text node count must be from 1 through 1000');
  }
  let characters = 0;
  const output = input.map((node) => {
    if (node === null || typeof node !== 'object' || typeof node.text !== 'string') {
      throw new TypeError('Comment rich-text node is invalid');
    }
    const text = [...node.text]
      .filter((character) => {
        const code = character.codePointAt(0) ?? 0;
        return code >= 32 || character === '\n' || character === '\t';
      })
      .join('');
    characters += [...text].length;
    if (characters > 20_000) throw new RangeError('Comment rich text exceeds 20000 characters');
    return Object.freeze({
      text,
      ...(node.bold === true ? { bold: true } : {}),
      ...(node.italic === true ? { italic: true } : {}),
      ...(node.code === true ? { code: true } : {}),
    });
  });
  return Object.freeze(output);
}

function snapshotAnchor(anchor: CommentAnchor): CommentAnchor {
  let snapshot: CommentAnchor;
  try {
    snapshot = JSON.parse(JSON.stringify(anchor)) as CommentAnchor;
  } catch {
    throw new TypeError('Comment anchor must be JSON serializable');
  }
  const location = snapshot.type === 'orphaned' ? snapshot.lastKnown : snapshot;
  if (location.type === 'cell') {
    if (
      !identifierPattern.test(location.cell.sheetId) ||
      !Number.isSafeInteger(location.cell.row) ||
      location.cell.row < 0 ||
      !Number.isSafeInteger(location.cell.column) ||
      location.cell.column < 0
    ) {
      throw new TypeError('Comment cell anchor is invalid');
    }
  } else if (location.type === 'range') {
    if (
      !identifierPattern.test(location.range.sheetId) ||
      !Number.isSafeInteger(location.range.start.row) ||
      !Number.isSafeInteger(location.range.start.column) ||
      !Number.isSafeInteger(location.range.end.row) ||
      !Number.isSafeInteger(location.range.end.column) ||
      location.range.start.row < 0 ||
      location.range.start.column < 0 ||
      location.range.end.row < location.range.start.row ||
      location.range.end.column < location.range.start.column
    ) {
      throw new TypeError('Comment range anchor is invalid');
    }
  } else if (
    !identifierPattern.test(location.sheetId) ||
    !identifierPattern.test(location.objectId)
  ) {
    throw new TypeError('Comment object anchor is invalid');
  }
  return deepFreeze(snapshot);
}

/** Creates a local comment model; persistence and remote synchronization remain host-owned. */
export function createCommentStore(options: {
  readonly documentId: string;
  readonly actorId: string;
  readonly permissions: () => PermissionSnapshot | undefined;
  readonly subscribePermissions?: (listener: () => void) => () => void;
  readonly nextId: () => string;
  readonly nextRevision: () => string;
}): CommentStore {
  if (!identifierPattern.test(options.documentId) || !identifierPattern.test(options.actorId)) {
    throw new TypeError('Comment store identity is invalid');
  }
  const threads = new Map<string, CommentThread>();
  const listeners = new Set<() => void>();
  let snapshot: readonly CommentThread[] = Object.freeze([]);
  const permission = (
    action: 'comment:view' | 'comment:create' | 'comment:resolve',
    threadId: string,
  ): void => {
    if (
      !(
        options.permissions()?.can(action, {
          type: 'comment',
          threadId,
        }) ?? false
      )
    ) {
      throw new TypeError(`Comment ${action} permission denied`);
    }
  };
  const current = (threadId: string, revision: string): CommentThread => {
    const thread = threads.get(threadId);
    if (thread === undefined) throw new TypeError('Comment thread does not exist');
    if (thread.revision !== revision) throw new TypeError('Comment thread revision conflict');
    return thread;
  };
  const refresh = (): void => {
    snapshot = Object.freeze(
      [...threads.values()].filter((thread) => {
        try {
          permission('comment:view', thread.id);
          return true;
        } catch {
          return false;
        }
      }),
    );
    for (const listener of listeners) listener();
  };
  const nextThread = (thread: Omit<CommentThread, 'revision'>): CommentThread => {
    const result = deepFreeze({
      ...thread,
      revision: options.nextRevision(),
    });
    if (!identifierPattern.test(result.revision)) {
      throw new TypeError('Comment thread revision is invalid');
    }
    threads.set(result.id, result);
    refresh();
    return result;
  };
  const store: CommentStore = {
    list(): readonly CommentThread[] {
      return snapshot;
    },
    getSnapshot(): readonly CommentThread[] {
      return snapshot;
    },
    subscribe(listener): () => void {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    refresh,
    create(input): CommentThread {
      if (input.expectedDocumentRevision !== input.currentDocumentRevision) {
        throw new TypeError('Comment document revision conflict');
      }
      const id = options.nextId();
      if (!identifierPattern.test(id) || threads.has(id)) {
        throw new TypeError('Comment thread ID is invalid or duplicated');
      }
      permission('comment:create', id);
      return nextThread({
        id,
        anchor: snapshotAnchor(input.anchor),
        messages: Object.freeze([
          deepFreeze({
            id: `${id}:message-1`,
            authorId: options.actorId,
            content: sanitizeCommentRichText(input.content),
          }),
        ]),
        resolved: false,
      });
    },
    reply(threadId, expectedThreadRevision, content): CommentThread {
      permission('comment:create', threadId);
      const thread = current(threadId, expectedThreadRevision);
      return nextThread({
        ...thread,
        messages: Object.freeze([
          ...thread.messages,
          deepFreeze({
            id: `${threadId}:message-${thread.messages.length + 1}`,
            authorId: options.actorId,
            content: sanitizeCommentRichText(content),
          }),
        ]),
      });
    },
    resolve(threadId, expectedThreadRevision, resolved): CommentThread {
      permission('comment:resolve', threadId);
      return nextThread({ ...current(threadId, expectedThreadRevision), resolved });
    },
    remove(threadId, expectedThreadRevision): void {
      permission('comment:resolve', threadId);
      current(threadId, expectedThreadRevision);
      threads.delete(threadId);
      refresh();
    },
    rebase(request): readonly CommentThread[] {
      const updates: CommentThread[] = [];
      for (const thread of threads.values()) {
        const anchor = transformCommentAnchor(thread.anchor, request);
        if (anchor === thread.anchor) continue;
        updates.push(nextThread({ ...thread, anchor }));
      }
      return Object.freeze(updates);
    },
  };
  options.subscribePermissions?.(refresh);
  return Object.freeze(store);
}

/** Projects comments into an explicit print policy without mutating thread state. */
export function projectCommentPrintContent(
  threads: readonly CommentThread[],
  policy: CommentPrintPolicy,
): readonly CommentPrintContent[] {
  if (policy === 'exclude') return Object.freeze([]);
  return Object.freeze(
    threads.map((thread, index) =>
      deepFreeze({
        threadId: thread.id,
        marker: String(index + 1),
        anchor: thread.anchor,
        ...(policy === 'full'
          ? {
              text: thread.messages
                .flatMap(({ content }) => content.map(({ text }) => text))
                .join('\n'),
            }
          : {}),
      }),
    ),
  );
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
  let operationTail: Promise<void> = Promise.resolve();
  const serialize = <Result>(operation: () => Promise<Result>): Promise<Result> => {
    const result = operationTail.then(operation);
    operationTail = result.then(
      () => undefined,
      () => undefined,
    );
    return result;
  };
  const validateStored = (
    documentId: string,
    stored: readonly CommentAnchorUpdateBatch[],
  ): readonly CommentAnchorUpdateBatch[] => {
    if (stored.length > 10_000) throw new RangeError('Comment outbox resume limit is 10000');
    const snapshots = stored.map(snapshotBatch);
    const operationIds = new Set<string>();
    for (const [index, item] of snapshots.entries()) {
      if (item.documentId !== documentId) {
        throw new TypeError('Comment outbox batch documentId does not match resume document');
      }
      if (
        item.fromDocumentRevision === item.toDocumentRevision ||
        operationIds.has(item.operationId)
      ) {
        throw new TypeError('Comment outbox revision chain or operation order is invalid');
      }
      const previous = snapshots[index - 1];
      if (previous !== undefined && previous.toDocumentRevision !== item.fromDocumentRevision) {
        throw new TypeError('Comment outbox revision chain is not contiguous');
      }
      operationIds.add(item.operationId);
    }
    return snapshots;
  };
  const submitStored = async (
    batch: CommentAnchorUpdateBatch,
    signal: AbortSignal,
  ): Promise<CommentAnchorUpdateAck> => {
    if (signal.aborted) throw new TypeError('Comment anchor update was cancelled');
    const acknowledgement = await options.adapter.submit(batch, signal);
    if ('status' in acknowledgement) {
      if (
        acknowledgement.status !== 'conflict' ||
        acknowledgement.operationId !== batch.operationId ||
        acknowledgement.expectedRevision !== batch.fromDocumentRevision ||
        !identifierPattern.test(acknowledgement.expectedRevision) ||
        !identifierPattern.test(acknowledgement.currentRevision)
      ) {
        throw new TypeError('Comment anchor update conflict response is invalid');
      }
      throw new CommentAnchorOutboxConflictError(
        batch.operationId,
        acknowledgement.expectedRevision,
        acknowledgement.currentRevision,
        acknowledgement,
      );
    }
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
      return serialize(async () => {
        if (signal.aborted) throw new TypeError('Comment anchor update was cancelled');
        const pending = validateStored(
          snapshot.documentId,
          await options.outbox.list(snapshot.documentId),
        );
        const previous = pending[pending.length - 1];
        const duplicate = pending.find((item) => item.operationId === snapshot.operationId);
        if (
          duplicate !== undefined ||
          (previous !== undefined && previous.toDocumentRevision !== snapshot.fromDocumentRevision)
        ) {
          throw new CommentAnchorOutboxConflictError(
            snapshot.operationId,
            snapshot.fromDocumentRevision,
            previous?.toDocumentRevision ??
              duplicate?.toDocumentRevision ??
              snapshot.fromDocumentRevision,
          );
        }
        await options.outbox.put(snapshot);
        return submitStored(snapshot, signal);
      });
    },
    async resume(
      documentId: string,
      signal: AbortSignal,
    ): Promise<readonly CommentAnchorUpdateAck[]> {
      if (!identifierPattern.test(documentId)) throw new TypeError('Comment documentId is invalid');
      return serialize(async () => {
        if (signal.aborted) throw new TypeError('Comment anchor update was cancelled');
        const snapshots = validateStored(documentId, await options.outbox.list(documentId));
        const acknowledgements: CommentAnchorUpdateAck[] = [];
        for (const batch of snapshots) {
          acknowledgements.push(await submitStored(batch, signal));
        }
        return Object.freeze(acknowledgements);
      });
    },
  });
}
