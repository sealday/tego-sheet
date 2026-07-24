import { describe, expect, it, vi } from 'vitest';
import {
  createCollaborationOutboxCoordinator,
  createCollaborationSession,
  createControllerRemoteTransactionBoundary,
  createPresenceStore,
  createRemoteOperationProcessor,
  type CollaborationOutboundOperation,
} from '../../../src/integrations/collaboration';
import type { SerializableTransactionEnvelope } from '../../../src/core/controller/spreadsheet-document-controller';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { testDocument } from '../../helpers/workbook-builders';

const transaction: SerializableTransactionEnvelope = {
  schemaVersion: 1,
  id: 'transaction-1',
  baseRevision: 0,
  commands: [],
};

describe('collaboration integration contract', () => {
  it('deduplicates operations and validates revision order before atomic apply', () => {
    const commit = vi.fn();
    const rollback = vi.fn();
    const prepare = vi.fn(() => ({ commit, rollback }));
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: { prepare },
    });
    const operation = {
      operationId: 'operation-1',
      actorId: 'actor-2',
      baseRevision: 'revision-1',
      revision: 'revision-2',
      transaction,
    };

    expect(processor.process(operation)).toEqual({
      status: 'applied',
      revision: 'revision-2',
    });
    expect(processor.process(operation)).toEqual({
      status: 'duplicate',
      operationId: 'operation-1',
    });
    expect(prepare).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
    expect(rollback).not.toHaveBeenCalled();
  });

  it('requests resync for gaps and never advances revision after rejected apply', () => {
    const prepare = vi.fn(() => undefined);
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: { prepare },
    });

    expect(
      processor.process({
        operationId: 'operation-gap',
        actorId: 'actor-2',
        baseRevision: 'revision-0',
        revision: 'revision-2',
        transaction,
      }),
    ).toEqual({
      status: 'resync-required',
      expectedRevision: 'revision-1',
      receivedBaseRevision: 'revision-0',
    });
    expect(prepare).not.toHaveBeenCalled();
    processor.resetAfterResync('revision-1');

    expect(
      processor.process({
        operationId: 'operation-rejected',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction,
      }),
    ).toEqual({ status: 'rejected', operationId: 'operation-rejected' });
    expect(processor.revision).toBe('revision-1');
  });

  it('rejects oversized operations before apply', () => {
    const prepare = vi.fn();
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      maximumOperationBytes: 128,
      permissionGate: () => true,
      transactionBoundary: { prepare },
    });

    expect(() =>
      processor.process({
        operationId: 'operation-large',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction: {
          ...transaction,
          metadata: { payload: 'x'.repeat(256) },
        },
      }),
    ).toThrow(/128/u);
    expect(prepare).not.toHaveBeenCalled();
  });

  it('caps operation size, defaults permission to deny, and rolls back a failed commit', () => {
    expect(() =>
      createRemoteOperationProcessor({
        initialRevision: 'revision-1',
        maximumOperationBytes: 1024 * 1024 + 1,
        transactionBoundary: { prepare: vi.fn() },
      }),
    ).toThrow(/1 MiB/u);
    expect(() =>
      createRemoteOperationProcessor({
        initialRevision: 'revision-1',
        maximumRememberedOperations: 9_999,
        transactionBoundary: { prepare: vi.fn() },
      }),
    ).toThrow(/10000/u);

    const deniedPrepare = vi.fn();
    const denied = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      transactionBoundary: { prepare: deniedPrepare },
    });
    expect(
      denied.process({
        operationId: 'operation-denied',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction,
      }),
    ).toEqual({ status: 'rejected', operationId: 'operation-denied' });
    expect(deniedPrepare).not.toHaveBeenCalled();

    const rollback = vi.fn();
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({
          commit: () => {
            throw new Error('partial host failure');
          },
          rollback,
        }),
      },
    });
    expect(
      processor.process({
        operationId: 'operation-failed',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction,
      }),
    ).toEqual({ status: 'rejected', operationId: 'operation-failed' });
    expect(rollback).toHaveBeenCalledTimes(1);
    expect(processor.revision).toBe('revision-1');
  });

  it('requires resync when both remote commit and rollback fail', () => {
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({
          commit: () => {
            throw new Error('partial host failure');
          },
          rollback: () => {
            throw new Error('rollback failure');
          },
        }),
      },
    });
    const operation = {
      operationId: 'operation-corrupt',
      actorId: 'actor-2',
      baseRevision: 'revision-1',
      revision: 'revision-2',
      transaction,
    };

    expect(processor.process(operation)).toEqual({
      status: 'resync-required',
      expectedRevision: 'revision-1',
      receivedBaseRevision: 'revision-1',
    });
    expect(
      processor.process({
        ...operation,
        operationId: 'operation-after-corrupt',
      }),
    ).toMatchObject({ status: 'resync-required' });
  });

  it('rejects local-history commands in favor of inverse transactions', () => {
    const prepare = vi.fn();
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: { prepare },
    });

    expect(() =>
      processor.process({
        operationId: 'operation-undo-command',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction: {
          ...transaction,
          commands: [
            {
              schemaVersion: 1,
              id: 'command-1',
              command: { type: 'undo' },
            },
          ],
        },
      }),
    ).toThrow(/inverse transaction/u);
    expect(prepare).not.toHaveBeenCalled();
  });

  it('blocks after disconnect or revision gap until an explicit reset', () => {
    const commit = vi.fn();
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({ commit, rollback: vi.fn() }),
      },
    });
    processor.disconnect();
    expect(
      processor.process({
        operationId: 'operation-offline',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction,
      }),
    ).toMatchObject({ status: 'disconnected' });
    processor.resetAfterResync('revision-1');
    expect(
      processor.process({
        operationId: 'operation-applied',
        actorId: 'actor-2',
        baseRevision: 'revision-1',
        revision: 'revision-2',
        transaction,
      }),
    ).toMatchObject({ status: 'applied' });
  });

  it('durably queues local operations and negotiates replay and collaborative undo', async () => {
    const stored = new Map<string, CollaborationOutboundOperation>();
    const outbox = {
      put: vi.fn(async (operation: CollaborationOutboundOperation) => {
        stored.set(operation.operationId, operation);
      }),
      remove: vi.fn(async (operationId: string) => {
        stored.delete(operationId);
      }),
      list: vi.fn(async () => [...stored.values()]),
    };
    const submit = vi.fn(async (operation: CollaborationOutboundOperation) => ({
      operationId: operation.operationId,
    }));
    const coordinator = createCollaborationOutboxCoordinator({
      documentId: 'document-1',
      outbox,
      adapter: { submit },
    });
    coordinator.disconnect();
    const inverse = {
      ...transaction,
      id: 'transaction-undo',
    };
    await expect(
      coordinator.queue(
        {
          protocolVersion: 1,
          documentId: 'document-1',
          operationId: 'operation-undo',
          baseRevision: 'revision-2',
          revision: 'revision-3',
          undoOfOperationId: 'operation-1',
          transaction: inverse,
        },
        new AbortController().signal,
      ),
    ).resolves.toEqual({ status: 'queued', operationId: 'operation-undo' });
    expect(submit).not.toHaveBeenCalled();

    await expect(
      coordinator.reconnect(
        { protocolVersions: [1], collaborativeUndo: false },
        new AbortController().signal,
      ),
    ).rejects.toThrow(/undo capability/u);
    expect(stored.size).toBe(1);
    await coordinator.reconnect(
      { protocolVersions: [1], collaborativeUndo: true },
      new AbortController().signal,
    );
    expect(submit).toHaveBeenCalledTimes(1);
    expect(stored.size).toBe(0);
  });

  it('serializes new local operations behind reconnect replay', async () => {
    const stored = new Map<string, CollaborationOutboundOperation>();
    let releaseReplay!: () => void;
    const replayBarrier = new Promise<void>((resolve) => {
      releaseReplay = resolve;
    });
    const outbox = {
      put: vi.fn(async (operation: CollaborationOutboundOperation) => {
        stored.set(operation.operationId, operation);
      }),
      remove: vi.fn(async (operationId: string) => {
        stored.delete(operationId);
      }),
      list: vi.fn(async () => [...stored.values()]),
    };
    stored.set('operation-old', {
      protocolVersion: 1,
      documentId: 'document-1',
      operationId: 'operation-old',
      baseRevision: 'revision-1',
      revision: 'revision-2',
      transaction,
    });
    const submit = vi.fn(async (operation: CollaborationOutboundOperation) => {
      if (operation.operationId === 'operation-old') await replayBarrier;
      return { operationId: operation.operationId };
    });
    const coordinator = createCollaborationOutboxCoordinator({
      documentId: 'document-1',
      outbox,
      adapter: { submit },
    });
    const reconnect = coordinator.reconnect(
      { protocolVersions: [1], collaborativeUndo: true },
      new AbortController().signal,
    );
    await vi.waitFor(() => expect(submit).toHaveBeenCalledTimes(1));
    const queued = coordinator.queue(
      {
        protocolVersion: 1,
        documentId: 'document-1',
        operationId: 'operation-new',
        baseRevision: 'revision-2',
        revision: 'revision-3',
        transaction: { ...transaction, id: 'transaction-new' },
      },
      new AbortController().signal,
    );
    await Promise.resolve();
    expect(submit).toHaveBeenCalledTimes(1);

    releaseReplay();
    await expect(reconnect).resolves.toEqual(['operation-old']);
    await expect(queued).resolves.toEqual({
      status: 'submitted',
      operationId: 'operation-new',
    });
    expect(submit.mock.calls.map(([operation]) => operation.operationId)).toEqual([
      'operation-old',
      'operation-new',
    ]);
  });

  it('stays offline after replay failure so new work cannot overtake the outbox', async () => {
    const stored = new Map<string, CollaborationOutboundOperation>();
    stored.set('operation-old', {
      protocolVersion: 1,
      documentId: 'document-1',
      operationId: 'operation-old',
      baseRevision: 'revision-1',
      revision: 'revision-2',
      transaction,
    });
    const outbox = {
      put: vi.fn(async (operation: CollaborationOutboundOperation) => {
        stored.set(operation.operationId, operation);
      }),
      remove: vi.fn(async (operationId: string) => {
        stored.delete(operationId);
      }),
      list: vi.fn(async () => [...stored.values()]),
    };
    const submit = vi.fn(async () => {
      throw new Error('offline');
    });
    const coordinator = createCollaborationOutboxCoordinator({
      documentId: 'document-1',
      outbox,
      adapter: { submit },
    });

    await expect(
      coordinator.reconnect(
        { protocolVersions: [1], collaborativeUndo: true },
        new AbortController().signal,
      ),
    ).rejects.toThrow(/offline/u);
    await expect(
      coordinator.queue(
        {
          protocolVersion: 1,
          documentId: 'document-1',
          operationId: 'operation-new',
          baseRevision: 'revision-2',
          revision: 'revision-3',
          transaction: { ...transaction, id: 'transaction-new' },
        },
        new AbortController().signal,
      ),
    ).resolves.toEqual({ status: 'queued', operationId: 'operation-new' });
    expect(submit).toHaveBeenCalledTimes(1);
  });

  it('keeps expiring presence in session state only and releases it on close', () => {
    let now = 100;
    const store = createPresenceStore({ now: () => now });
    store.replace([
      {
        actorId: 'actor-2',
        sheetId: 'sheet-1',
        selections: [],
        display: { label: 'Remote\u0000 User', color: '#ff0000' },
        expiresAt: 200,
      },
    ]);

    expect(store.list()[0]?.display.label).toBe('Remote User');
    now = 201;
    expect(store.list()).toEqual([]);
    store.close();
    expect(() => store.replace([])).toThrow(/closed/u);
  });

  it('publishes presence expiry without requiring a workbook or UI read', () => {
    vi.useFakeTimers();
    try {
      vi.setSystemTime(0);
      const listener = vi.fn();
      const store = createPresenceStore();
      store.subscribe(listener);
      store.replace([
        {
          actorId: 'actor-2',
          sheetId: 'sheet-1',
          selections: [],
          display: { label: 'Remote', color: '#ff0000' },
          expiresAt: 100,
        },
      ]);

      vi.advanceTimersByTime(101);

      expect(store.getSnapshot()).toEqual([]);
      expect(listener).toHaveBeenCalledTimes(2);
      store.close();
    } finally {
      vi.useRealTimers();
    }
  });

  it('commits remote transactions through the reversible production controller boundary', () => {
    const document = testDocument([{ name: 'A' }]);
    const controller = new SpreadsheetDocumentController(document);
    const sheet = controller.getSheetIds()[0]!;
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      permissionGate: () => true,
      transactionBoundary: createControllerRemoteTransactionBoundary(controller),
    });
    const result = processor.process({
      operationId: 'operation-controller',
      actorId: 'actor-2',
      baseRevision: 'revision-1',
      revision: 'revision-2',
      transaction: {
        schemaVersion: 1,
        id: 'transaction-controller',
        baseRevision: 0,
        commands: [
          {
            schemaVersion: 1,
            id: 'command-controller',
            command: {
              type: 'set-cell-text',
              address: { sheet, row: 0, column: 0 },
              text: 'remote',
            },
          },
        ],
      },
    });

    expect(result).toEqual({ status: 'applied', revision: 'revision-2' });
    expect(controller.getCellText({ sheet, row: 0, column: 0 })).toBe('remote');
    controller.dispose();
  });

  it('connects host subscriptions and routes operation/presence events', async () => {
    let listener:
      | ((
          event: import('../../../src/integrations/collaboration').CollaborationInboundEvent,
        ) => void)
      | undefined;
    const presence = createPresenceStore({ now: () => 0 });
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-0',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({ commit: vi.fn(), rollback: vi.fn() }),
      },
    });
    const session = createCollaborationSession({
      processor,
      presence,
      port: {
        connect: async () => ({
          revision: 'revision-1',
          capabilities: { protocolVersions: [1], collaborativeUndo: true },
        }),
        subscribe: (next) => {
          listener = next;
          return () => {
            listener = undefined;
          };
        },
      },
    });
    await session.connect(new AbortController().signal);
    listener?.({
      type: 'presence',
      presence: [
        {
          actorId: 'actor-2',
          sheetId: 'sheet-1',
          selections: [],
          display: { label: 'Remote', color: '#ff0000' },
          expiresAt: 100,
        },
      ],
    });
    expect(presence.getSnapshot()).toHaveLength(1);
    session.disconnect();
    expect(presence.getSnapshot()).toEqual([]);
    expect(session.state.status).toBe('disconnected');
  });

  it('returns to disconnected after a failed handshake and requests resync for bad events', async () => {
    const presence = createPresenceStore({ now: () => 0 });
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-0',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({ commit: vi.fn(), rollback: vi.fn() }),
      },
    });
    const failed = createCollaborationSession({
      processor,
      presence,
      port: {
        connect: async () => {
          throw new Error('offline');
        },
        subscribe: vi.fn(),
      },
    });
    await expect(failed.connect(new AbortController().signal)).rejects.toThrow(/offline/u);
    expect(failed.getSnapshot()).toEqual({ status: 'disconnected' });

    let listener:
      | ((
          event: import('../../../src/integrations/collaboration').CollaborationInboundEvent,
        ) => void)
      | undefined;
    const session = createCollaborationSession({
      processor,
      presence,
      port: {
        connect: async () => ({
          revision: 'revision-1',
          capabilities: { protocolVersions: [1], collaborativeUndo: true },
        }),
        subscribe: (next) => {
          listener = next;
          return () => undefined;
        },
      },
    });
    await session.connect(new AbortController().signal);

    expect(() =>
      listener?.({
        type: 'presence',
        presence: [
          {
            actorId: 'actor-2',
            sheetId: 'sheet-1',
            selections: [],
            display: { label: 'Remote', color: 'invalid' },
            expiresAt: 100,
          },
        ],
      }),
    ).not.toThrow();
    expect(session.getSnapshot()).toEqual({
      status: 'resync-required',
      revision: 'revision-1',
    });
  });

  it('does not let an obsolete handshake failure disconnect a newer connection', async () => {
    let rejectFirst!: (error: Error) => void;
    const firstHandshake = new Promise<never>((_resolve, reject) => {
      rejectFirst = reject;
    });
    let attempt = 0;
    const session = createCollaborationSession({
      processor: createRemoteOperationProcessor({
        initialRevision: 'revision-0',
        permissionGate: () => true,
        transactionBoundary: {
          prepare: () => ({ commit: vi.fn(), rollback: vi.fn() }),
        },
      }),
      presence: createPresenceStore({ now: () => 0 }),
      port: {
        connect: async () => {
          attempt += 1;
          if (attempt === 1) return firstHandshake;
          return {
            revision: 'revision-2',
            capabilities: { protocolVersions: [1], collaborativeUndo: true },
          };
        },
        subscribe: () => () => undefined,
      },
    });

    const obsolete = session.connect(new AbortController().signal);
    await vi.waitFor(() => expect(session.state.status).toBe('connecting'));
    await session.connect(new AbortController().signal);
    expect(session.getSnapshot()).toEqual({ status: 'connected', revision: 'revision-2' });

    rejectFirst(new Error('obsolete offline'));
    await expect(obsolete).rejects.toThrow(/obsolete offline/u);
    expect(session.getSnapshot()).toEqual({ status: 'connected', revision: 'revision-2' });
  });

  it('cancels an in-flight handshake and remains disconnected', async () => {
    const controller = new AbortController();
    const session = createCollaborationSession({
      processor: createRemoteOperationProcessor({
        initialRevision: 'revision-0',
        permissionGate: () => true,
        transactionBoundary: {
          prepare: () => ({ commit: vi.fn(), rollback: vi.fn() }),
        },
      }),
      presence: createPresenceStore({ now: () => 0 }),
      port: {
        connect: (signal) =>
          new Promise((_resolve, reject) => {
            signal.addEventListener('abort', () => reject(new TypeError('handshake aborted')), {
              once: true,
            });
          }),
        subscribe: vi.fn(),
      },
    });

    const connecting = session.connect(controller.signal);
    controller.abort();
    await expect(connecting).rejects.toThrow(/aborted|cancelled/u);
    expect(session.getSnapshot()).toEqual({ status: 'disconnected' });
  });
});
