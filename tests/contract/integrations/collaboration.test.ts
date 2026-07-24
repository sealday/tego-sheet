import { describe, expect, it, vi } from 'vitest';
import {
  createCollaborationOutboxCoordinator,
  createPresenceStore,
  createRemoteOperationProcessor,
  type CollaborationOutboundOperation,
} from '../../../src/integrations/collaboration';
import type { SerializableTransactionEnvelope } from '../../../src/core/controller/spreadsheet-document-controller';

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
});
