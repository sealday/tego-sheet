import { describe, expect, it, vi } from 'vitest';
import { CoordinateTransform } from '../../../src/core/coordinates/coordinate-transform';
import {
  createCommentAnchorOutboxCoordinator,
  transformCommentAnchor,
  type CommentAnchorUpdateBatch,
} from '../../../src/integrations/comments';

const batch: CommentAnchorUpdateBatch = {
  operationId: 'operation-1',
  documentId: 'document-1',
  fromDocumentRevision: 'revision-1',
  toDocumentRevision: 'revision-2',
  updates: [
    {
      threadId: 'thread-1',
      expectedThreadRevision: 'thread-revision-1',
      anchor: {
        type: 'cell',
        cell: { sheetId: 'sheet-1', row: 4, column: 2 },
      },
    },
  ],
};

describe('comments integration contract', () => {
  it('transforms cell/range anchors and orphans deleted targets', () => {
    expect(
      transformCommentAnchor(
        { type: 'cell', cell: { sheetId: 'sheet-1', row: 4, column: 2 } },
        {
          sheetId: 'sheet-1',
          transform: CoordinateTransform.insert('row', 2, 3),
        },
      ),
    ).toEqual({ type: 'cell', cell: { sheetId: 'sheet-1', row: 7, column: 2 } });
    expect(
      transformCommentAnchor(
        {
          type: 'range',
          range: {
            sheetId: 'sheet-1',
            start: { row: 2, column: 0 },
            end: { row: 3, column: 1 },
          },
        },
        {
          sheetId: 'sheet-1',
          transform: CoordinateTransform.delete('row', 2, 2),
        },
      ),
    ).toEqual({
      type: 'orphaned',
      lastKnown: {
        type: 'range',
        range: {
          sheetId: 'sheet-1',
          start: { row: 2, column: 0 },
          end: { row: 3, column: 1 },
        },
      },
      reason: 'anchor target was deleted by a structural transaction',
    });
  });

  it('durably records a batch before remote submit and removes it only after ack', async () => {
    const events: string[] = [];
    const outbox = {
      put: vi.fn(async () => {
        events.push('put');
      }),
      remove: vi.fn(async () => {
        events.push('remove');
      }),
      list: vi.fn(async () => []),
    };
    const adapter = {
      submit: vi.fn(async () => {
        events.push('submit');
        return { operationId: 'operation-1' };
      }),
    };
    const coordinator = createCommentAnchorOutboxCoordinator({ outbox, adapter });

    await coordinator.queue(batch, new AbortController().signal);

    expect(events).toEqual(['put', 'submit', 'remove']);
  });

  it('retains failed batches and resumes with the same operation id', async () => {
    const stored = new Map<string, CommentAnchorUpdateBatch>();
    const outbox = {
      put: vi.fn(async (item: CommentAnchorUpdateBatch) => {
        stored.set(item.operationId, item);
      }),
      remove: vi.fn(async (operationId: string) => {
        stored.delete(operationId);
      }),
      list: vi.fn(async () => [...stored.values()]),
    };
    const submit = vi
      .fn()
      .mockRejectedValueOnce(new Error('offline'))
      .mockResolvedValueOnce({ operationId: 'operation-1' });
    const coordinator = createCommentAnchorOutboxCoordinator({
      outbox,
      adapter: { submit },
    });

    await expect(coordinator.queue(batch, new AbortController().signal)).rejects.toThrow('offline');
    expect(stored.has('operation-1')).toBe(true);
    await coordinator.resume('document-1', new AbortController().signal);

    expect(submit.mock.calls.map(([item]) => item.operationId)).toEqual([
      'operation-1',
      'operation-1',
    ]);
    expect(stored.size).toBe(0);
  });

  it('validates the entire document revision chain before resuming any batch', async () => {
    const submit = vi.fn(async (item: CommentAnchorUpdateBatch) => ({
      operationId: item.operationId,
    }));
    const coordinator = createCommentAnchorOutboxCoordinator({
      outbox: {
        put: vi.fn(),
        remove: vi.fn(),
        list: vi.fn(async () => [
          batch,
          {
            ...batch,
            operationId: 'operation-2',
            documentId: 'document-other',
            fromDocumentRevision: 'revision-wrong',
            toDocumentRevision: 'revision-3',
          },
        ]),
      },
      adapter: { submit },
    });

    await expect(coordinator.resume('document-1', new AbortController().signal)).rejects.toThrow(
      /documentId|revision chain/u,
    );
    expect(submit).not.toHaveBeenCalled();
  });
});
