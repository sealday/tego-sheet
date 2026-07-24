import { describe, expect, it, vi } from 'vitest';
import {
  createPresenceStore,
  createRemoteOperationProcessor,
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
    const apply = vi.fn(() => true);
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      apply,
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
    expect(apply).toHaveBeenCalledTimes(1);
  });

  it('requests resync for gaps and never advances revision after rejected apply', () => {
    const apply = vi.fn(() => false);
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      apply,
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
    expect(apply).not.toHaveBeenCalled();

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
    const apply = vi.fn(() => true);
    const processor = createRemoteOperationProcessor({
      initialRevision: 'revision-1',
      maximumOperationBytes: 128,
      apply,
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
    expect(apply).not.toHaveBeenCalled();
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
