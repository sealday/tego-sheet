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
  it('sends the largest request prefix that fits the byte budget', async () => {
    const requests: SaveRequest[] = [];
    const sampleRequest = (transactions: readonly SerializableTransactionEnvelope[]) => ({
      documentId: 'document-1',
      requestId: 'request-1',
      baseRevision: 'revision-1',
      transactions,
      reason: 'manual' as const,
    });
    const first = transaction('tx-1');
    const second = transaction('tx-2');
    const oneTransactionBytes = new TextEncoder().encode(
      JSON.stringify(sampleRequest([first])),
    ).byteLength;
    const twoTransactionBytes = new TextEncoder().encode(
      JSON.stringify(sampleRequest([first, second])),
    ).byteLength;
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: {
        save: async (request) => {
          requests.push(request);
          return {
            status: 'saved',
            revision: `revision-${requests.length + 1}`,
            persistedTransactionIds: request.transactions.map(({ id }) => id),
          };
        },
      },
      requestId: () => 'request-1',
      maximumRequestBytes: twoTransactionBytes - 1,
    });
    expect(oneTransactionBytes).toBeLessThan(twoTransactionBytes - 1);
    controller.enqueue(first);
    controller.enqueue(second);

    await controller.save();

    expect(requests[0]?.transactions.map(({ id }) => id)).toEqual(['tx-1']);
    expect(controller.state).toMatchObject({ status: 'dirty', pending: ['tx-2'] });
  });

  it('rejects only when the first pending transaction cannot fit the byte budget', async () => {
    const controller = createPersistenceController({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: { save: vi.fn() },
      requestId: () => 'request-1',
      maximumRequestBytes: 1,
    });
    controller.enqueue(transaction('tx-1'));

    await expect(controller.save()).rejects.toThrow(/exceeds/u);
  });

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
      autosaveDelayMs: 250,
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

  it('rejects controllers whose document identity does not match the persistence session', () => {
    const session = createPersistenceSession({
      documentId: 'document-1',
      initialRevision: 'revision-1',
      adapter: { save: vi.fn() },
      requestId: () => 'request-1',
    });
    const controller = new SpreadsheetDocumentController(testDocument([{ name: 'A' }]));

    expect(() => session.attachController(controller)).toThrow(/document.*identity/u);
    session.dispose();
    controller.dispose();
  });

  it('reschedules autosave when edits arrive during an in-flight batch', async () => {
    const pending = deferred<SaveResult>();
    const timers: Array<() => void> = [];
    const document = testDocument([{ name: 'A' }]);
    const save = vi
      .fn<(request: SaveRequest, signal: AbortSignal) => Promise<SaveResult>>()
      .mockImplementationOnce(() => pending.promise)
      .mockImplementationOnce(async (request) => ({
        status: 'saved',
        revision: 'revision-3',
        persistedTransactionIds: request.transactions.map(({ id }) => id),
      }));
    const session = createPersistenceSession({
      documentId: document.id,
      initialRevision: 'revision-1',
      adapter: { save },
      requestId: () => `request-${save.mock.calls.length + 1}`,
      autosaveDelayMs: 250,
      setTimer: (callback) => {
        timers.push(callback);
        return callback;
      },
      clearTimer: (handle) => {
        const index = timers.indexOf(handle as () => void);
        if (index >= 0) timers.splice(index, 1);
      },
    });
    const controller = new SpreadsheetDocumentController(document);
    session.attachController(controller);
    const sheet = controller.getSheetIds()[0]!;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'first' },
      'ref',
    );
    timers.shift()?.();
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 1 }, text: 'second' },
      'ref',
    );
    timers.shift()?.();
    pending.resolve({
      status: 'saved',
      revision: 'revision-2',
      persistedTransactionIds: [save.mock.calls[0]![0].transactions[0]!.id],
    });
    await pending.promise;
    await vi.waitFor(() => expect(timers).toHaveLength(1));
    timers.shift()?.();
    await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    await vi.waitFor(() => expect(session.state.status).toBe('clean'));
    session.dispose();
    controller.dispose();
  });

  it('bounds debounce, enforces max-wait, retries the same request, and cleans unload bindings', async () => {
    expect(() =>
      createPersistenceSession({
        documentId: 'document-1',
        initialRevision: 'revision-1',
        adapter: { save: vi.fn() },
        requestId: () => 'request-1',
        autosaveDelayMs: 249,
      }),
    ).toThrow(/250/u);

    let now = 0;
    const document = testDocument([{ name: 'A' }]);
    const delays: number[] = [];
    const listeners = new Set<(event: BeforeUnloadEvent) => void>();
    const save = vi
      .fn<(request: SaveRequest, signal: AbortSignal) => Promise<SaveResult>>()
      .mockRejectedValueOnce(new Error('offline'))
      .mockImplementationOnce(async (request) => ({
        status: 'saved',
        revision: 'revision-2',
        persistedTransactionIds: request.transactions.map(({ id }) => id),
      }));
    const session = createPersistenceSession({
      documentId: document.id,
      initialRevision: 'revision-1',
      adapter: { save },
      requestId: () => 'stable-request',
      autosaveDelayMs: 1_000,
      autosaveMaxWaitMs: 10_000,
      now: () => now,
      setTimer: (_callback, delay) => {
        delays.push(delay);
        return delay;
      },
      clearTimer: () => undefined,
    });
    const controller = new SpreadsheetDocumentController(document);
    session.attachController(controller);
    const sheet = controller.getSheetIds()[0]!;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'first' },
      'ref',
    );
    now = 9_750;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 1 }, text: 'second' },
      'ref',
    );
    expect(delays.at(-1)).toBe(250);
    await expect(session.save()).rejects.toThrow(/failed/u);
    await expect(session.retry()).resolves.toMatchObject({ status: 'saved' });
    expect(save.mock.calls.map(([request]) => request.requestId)).toEqual([
      'stable-request',
      'stable-request',
    ]);
    session.bindBeforeUnload({
      addEventListener: (_type, listener) => listeners.add(listener),
      removeEventListener: (_type, listener) => listeners.delete(listener),
    });
    expect(listeners.size).toBe(1);
    session.dispose();
    expect(listeners.size).toBe(0);
    controller.dispose();
  });

  it.each(['conflict', 'rejected', 'error'] as const)(
    'stops automatic retry after an autosave %s',
    async (outcome) => {
      const timers: Array<() => void> = [];
      const document = testDocument([{ name: 'A' }]);
      const adapter = {
        save: vi.fn(async (): Promise<SaveResult> => {
          if (outcome === 'conflict') {
            return { status: 'conflict', currentRevision: 'revision-remote' };
          }
          if (outcome === 'rejected') {
            return { status: 'rejected', code: 'DENIED', message: 'denied' };
          }
          throw new Error('offline');
        }),
      };
      const session = createPersistenceSession({
        documentId: document.id,
        initialRevision: 'revision-1',
        adapter,
        requestId: () => 'request-1',
        autosaveDelayMs: 250,
        setTimer: (callback) => {
          timers.push(callback);
          return callback;
        },
        clearTimer: (handle) => {
          const index = timers.indexOf(handle as () => void);
          if (index >= 0) timers.splice(index, 1);
        },
      });
      const controller = new SpreadsheetDocumentController(document);
      session.attachController(controller);
      const sheet = controller.getSheetIds()[0]!;
      controller.dispatch(
        { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: outcome },
        'ref',
      );
      timers.shift()?.();
      await vi.waitFor(() =>
        expect(session.state.status).toBe(outcome === 'conflict' ? 'conflict' : 'error'),
      );
      expect(timers).toEqual([]);
      expect(adapter.save).toHaveBeenCalledTimes(1);
      session.dispose();
      controller.dispose();
    },
  );

  it('resolves keep-local conflicts only against the expected remote revision and uses a new request', async () => {
    const requests: SaveRequest[] = [];
    let requestNumber = 0;
    const document = testDocument([{ name: 'A' }]);
    const session = createPersistenceSession({
      documentId: document.id,
      initialRevision: 'revision-1',
      adapter: {
        save: async (request) => {
          requests.push(request);
          if (requests.length === 1) {
            return { status: 'conflict', currentRevision: 'revision-remote' };
          }
          return {
            status: 'saved',
            revision: 'revision-final',
            persistedTransactionIds: request.transactions.map(({ id }) => id),
          };
        },
      },
      requestId: () => `request-${++requestNumber}`,
      autosaveDelayMs: 60_000,
    });
    const controller = new SpreadsheetDocumentController(document);
    session.attachController(controller);
    const sheet = controller.getSheetIds()[0]!;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'local' },
      'ref',
    );
    await session.save();
    expect(() =>
      session.resolveConflict({
        strategy: 'keep-local',
        expectedRemoteRevision: 'revision-stale',
      }),
    ).toThrow(/stale/u);
    expect(session.state.status).toBe('conflict');

    session.resolveConflict({
      strategy: 'keep-local',
      expectedRemoteRevision: 'revision-remote',
    });
    await session.save();
    expect(requests.map(({ requestId, baseRevision }) => ({ requestId, baseRevision }))).toEqual([
      { requestId: 'request-1', baseRevision: 'revision-1' },
      { requestId: 'request-2', baseRevision: 'revision-remote' },
    ]);
    session.dispose();
    controller.dispose();
  });

  it('validates load/merge conflict inputs before atomic controller replacement', async () => {
    const document = testDocument([{ name: 'A' }]);
    const session = createPersistenceSession({
      documentId: document.id,
      initialRevision: 'revision-1',
      adapter: {
        save: async () => ({ status: 'conflict', currentRevision: 'revision-remote' }),
      },
      requestId: () => 'request-1',
      autosaveDelayMs: 60_000,
    });
    const controller = new SpreadsheetDocumentController(document);
    session.attachController(controller);
    const sheet = controller.getSheetIds()[0]!;
    controller.dispatch(
      { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'local' },
      'ref',
    );
    await session.save();
    const before = controller.getDocument();

    expect(() =>
      session.resolveConflict({
        strategy: 'load-remote',
        expectedRemoteRevision: 'revision-remote',
        document: { ...document, schemaVersion: 999 } as never,
      }),
    ).toThrow(/document is invalid/u);
    expect(controller.getDocument()).toEqual(before);
    expect(session.state.status).toBe('conflict');
    expect(() =>
      session.resolveConflict({
        strategy: 'merge',
        expectedRemoteRevision: 'revision-remote',
        document,
        rebasedTransactions: [{ ...transaction('rebased'), schemaVersion: 2 } as never],
      }),
    ).toThrow(/transaction is invalid/u);
    expect(controller.getDocument()).toEqual(before);
    expect(session.state.status).toBe('conflict');

    const remote = testDocument([{ name: 'Remote' }]);
    session.resolveConflict({
      strategy: 'load-remote',
      expectedRemoteRevision: 'revision-remote',
      document: remote,
    });
    expect(session.state.status).toBe('clean');
    expect(controller.getDocument().workbook.sheets[0]?.name).toBe('Remote');
    session.dispose();
    controller.dispose();
  });
});
