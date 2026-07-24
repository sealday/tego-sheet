import { describe, expect, it, vi } from 'vitest';
import {
  createPersistenceSession,
  createPersistenceController,
  loadPersistedDocument,
  type SaveRequest,
  type SaveResult,
} from '../../../src/integrations/persistence';
import type { SerializableTransactionEnvelope } from '../../../src/core/controller/spreadsheet-document-controller';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { testDocument } from '../../helpers/workbook-builders';

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

  it('sanitizes explicit adapter rejections before publishing them', async () => {
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: {
        save: async () => ({
          status: 'rejected',
          code: 'TOKEN_LEAK',
          message: 'Bearer secret-value\n    at private/server.ts:1',
        }),
      },
      requestId: () => 'request-1',
    });
    controller.enqueue(transaction('tx-1'));

    await expect(controller.save()).resolves.toEqual({
      status: 'rejected',
      code: 'PERSISTENCE_SAVE_REJECTED',
      message: 'Persistence adapter rejected the save',
    });
    expect(JSON.stringify(controller.state)).not.toContain('secret-value');
    expect(JSON.stringify(controller.state)).not.toContain('private/server');
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

  it('validates a complete loaded document envelope before exposure', async () => {
    const document = testDocument([{ name: 'A' }]);
    await expect(
      loadPersistedDocument(
        {
          load: async () => ({
            schemaVersion: 1,
            documentId: document.id,
            revision: 'revision-1',
            document,
          }),
        },
        document.id,
        new AbortController().signal,
      ),
    ).resolves.toMatchObject({ revision: 'revision-1', document });
    await expect(
      loadPersistedDocument(
        {
          load: async () => ({
            schemaVersion: 1,
            documentId: document.id,
            revision: 'revision-1',
            document: { ...document, schemaVersion: 999 },
          }),
        },
        document.id,
        new AbortController().signal,
      ),
    ).rejects.toThrow(/loaded document is invalid/u);
  });

  it('tracks controller commits, stays durable offline, and guards beforeunload', async () => {
    const document = testDocument([{ name: 'A' }]);
    const requests: SaveRequest[] = [];
    const session = createPersistenceSession({
      documentId: document.id,
      initialRevision: 'revision-1',
      adapter: {
        save: async (request) => {
          requests.push(request);
          return {
            status: 'saved',
            revision: 'revision-2',
            persistedTransactionIds: request.transactions.map(({ id }) => id),
          };
        },
      },
      requestId: () => 'request-1',
      autosaveDelayMs: 10,
    });
    const controller = new SpreadsheetDocumentController(document);
    session.attachController(controller);
    const sheet = controller.getSheetIds()[0]!;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'local' },
      'ref',
    );
    session.setOnline(false);

    expect(session.state).toMatchObject({ status: 'offline' });
    await expect(session.save()).resolves.toEqual({ status: 'offline' });
    expect(requests).toEqual([]);

    let beforeUnload: ((event: BeforeUnloadEvent) => void) | undefined;
    const unbind = session.bindBeforeUnload({
      addEventListener: (_type, listener) => {
        beforeUnload = listener;
      },
      removeEventListener: () => undefined,
    });
    const event = {
      preventDefault: vi.fn(),
      returnValue: undefined,
    } as unknown as BeforeUnloadEvent;
    beforeUnload?.(event);
    expect(event.preventDefault).toHaveBeenCalled();
    expect(event.returnValue).toBe('');

    session.setOnline(true);
    await expect(session.save('before-close')).resolves.toMatchObject({ status: 'saved' });
    expect(requests).toHaveLength(1);
    unbind();
    session.dispose();
    controller.dispose();
  });
});
