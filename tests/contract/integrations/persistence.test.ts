import { describe, expect, it, vi } from 'vitest';
import {
  createPersistenceController,
  type SaveRequest,
  type SaveResult,
} from '../../../src/integrations/persistence';
import type { SerializableTransactionEnvelope } from '../../../src/core/controller/spreadsheet-document-controller';

function transaction(id: string): SerializableTransactionEnvelope {
  return {
    schemaVersion: 1,
    id,
    baseRevision: 0,
    commands: [],
  };
}

function deferred<Result>() {
  let resolve!: (value: Result) => void;
  const promise = new Promise<Result>((resolve_) => {
    resolve = resolve_;
  });
  return { promise, resolve };
}

describe('persistence integration contract', () => {
  it('keeps edits committed during a save in the next pending batch', async () => {
    const pending = deferred<SaveResult>();
    const save = vi.fn(() => pending.promise);
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: { save },
      requestId: () => 'request-1',
    });
    controller.enqueue(transaction('tx-1'));

    const saving = controller.save('manual');
    expect(controller.state).toEqual({
      status: 'saving',
      revision: 'revision-1',
      requestId: 'request-1',
      inFlight: ['tx-1'],
      pending: [],
    });
    controller.enqueue(transaction('tx-2'));
    pending.resolve({
      status: 'saved',
      revision: 'revision-2',
      persistedTransactionIds: ['tx-1'],
    });
    await saving;

    expect(controller.state).toEqual({
      status: 'dirty',
      revision: 'revision-2',
      pending: ['tx-2'],
    });
  });

  it('retries a failed request with the same idempotency key', async () => {
    const save = vi
      .fn<(request: SaveRequest, signal: AbortSignal) => Promise<SaveResult>>()
      .mockRejectedValueOnce(new Error('offline'))
      .mockResolvedValueOnce({
        status: 'saved',
        revision: 'revision-2',
        persistedTransactionIds: ['tx-1'],
      });
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: { save },
      requestId: () => 'stable-request',
    });
    controller.enqueue(transaction('tx-1'));

    await expect(controller.save()).rejects.toThrow('Persistence save failed');
    expect(controller.state.status).toBe('error');
    await expect(controller.retry()).resolves.toMatchObject({ status: 'saved' });
    expect(save.mock.calls[0]?.[0].requestId).toBe('stable-request');
    expect(save.mock.calls[1]?.[0].requestId).toBe('stable-request');
  });

  it('retains every unacknowledged transaction on conflict', async () => {
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: {
        save: async () => ({ status: 'conflict', currentRevision: 'revision-remote' }),
      },
      requestId: () => 'request-1',
    });
    controller.enqueue(transaction('tx-1'));

    await controller.save();

    expect(controller.state).toEqual({
      status: 'conflict',
      baseRevision: 'revision-1',
      currentRevision: 'revision-remote',
      pending: ['tx-1'],
    });
    await expect(controller.save()).rejects.toThrow(/resolve/u);
    controller.enqueue(transaction('tx-2'));
    expect(controller.state.status).toBe('conflict');

    controller.resolveConflict('revision-remote');
    expect(controller.state).toEqual({
      status: 'dirty',
      revision: 'revision-remote',
      pending: ['tx-1', 'tx-2'],
    });
  });

  it('rejects incomplete saved acknowledgements without advancing revision or pending state', async () => {
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: {
        save: async () => ({
          status: 'saved',
          revision: 'revision-2',
          persistedTransactionIds: [],
        }),
      },
      requestId: () => 'request-1',
    });
    controller.enqueue(transaction('tx-1'));

    await expect(controller.save()).rejects.toThrow(/complete in-flight batch/u);
    expect(controller.state).toMatchObject({
      status: 'error',
      revision: 'revision-1',
      pending: ['tx-1'],
    });
  });

  it('cancels in-flight work on disposal and rejects future saves', async () => {
    const save = vi.fn(
      (_request: SaveRequest, signal: AbortSignal) =>
        new Promise<SaveResult>((_resolve, reject) => {
          signal.addEventListener('abort', () => reject(new Error('cancelled')), { once: true });
        }),
    );
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: { save },
      requestId: () => 'request-1',
    });
    controller.enqueue(transaction('tx-1'));
    const saving = controller.save();
    controller.dispose();

    await expect(saving).rejects.toThrow('Persistence save was cancelled');
    await expect(controller.save()).rejects.toThrow(/disposed/u);
  });
});
