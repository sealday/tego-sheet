import { describe, expect, it, vi } from 'vitest';
import {
  createPermissionSnapshot,
  createPermissionStore,
  evaluatePermission,
} from '../../../src/integrations/permission';

describe('permission integration contract', () => {
  it('defaults protected actions to deny until an atomic snapshot is installed', () => {
    const store = createPermissionStore();
    expect(store.can('document:edit', { type: 'document', documentId: 'document-1' })).toBe(false);

    store.replace(
      createPermissionSnapshot({
        revision: 'permission-1',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:edit',
            target: { type: 'document', documentId: 'document-1' },
          },
        ],
      }),
    );

    expect(store.can('document:edit', { type: 'document', documentId: 'document-1' })).toBe(true);
    expect(store.can('download', { type: 'document', documentId: 'document-1' })).toBe(false);
  });

  it('denies a multi-target operation atomically when any target is denied', () => {
    const snapshot = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [
        {
          action: 'range:edit',
          target: {
            type: 'range',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 9, column: 9 },
            },
          },
        },
      ],
    });

    expect(
      evaluatePermission(snapshot, [
        {
          action: 'range:edit',
          target: {
            type: 'range',
            range: {
              sheetId: 'sheet-1',
              start: { row: 1, column: 1 },
              end: { row: 2, column: 2 },
            },
          },
        },
        {
          action: 'range:edit',
          target: {
            type: 'range',
            range: {
              sheetId: 'sheet-1',
              start: { row: 20, column: 0 },
              end: { row: 20, column: 0 },
            },
          },
        },
      ]),
    ).toEqual({
      allowed: false,
      snapshotRevision: 'permission-1',
      deniedTargets: [
        {
          type: 'range',
          range: {
            sheetId: 'sheet-1',
            start: { row: 20, column: 0 },
            end: { row: 20, column: 0 },
          },
        },
      ],
    });
  });

  it('rejects stale permission revisions and target explosions', () => {
    const snapshot = createPermissionSnapshot({
      revision: 'permission-2',
      actorId: 'actor-1',
      grants: [],
    });
    expect(() => evaluatePermission(snapshot, [], { expectedRevision: 'permission-1' })).toThrow(
      /stale/u,
    );
    expect(() =>
      evaluatePermission(
        snapshot,
        Array.from({ length: 10_001 }, () => ({
          action: 'print' as const,
          target: { type: 'document' as const, documentId: 'document-1' },
        })),
      ),
    ).toThrow(/10000/u);
  });

  it('publishes complete snapshot replacement after the new snapshot is visible', () => {
    const store = createPermissionStore();
    const observed: string[] = [];
    const listener = vi.fn((snapshot) => {
      observed.push(snapshot.revision);
      expect(store.getSnapshot()).toBe(snapshot);
    });
    const unsubscribe = store.subscribe(listener);
    const snapshot = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [],
    });

    store.replace(snapshot);
    unsubscribe();
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-2',
        actorId: 'actor-1',
        grants: [],
      }),
    );

    expect(observed).toEqual(['permission-1']);
  });

  it('publishes permission removal after clear is visible', () => {
    const store = createPermissionStore();
    const observed: Array<string | undefined> = [];
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-1',
        actorId: 'actor-1',
        grants: [],
      }),
    );
    store.subscribe((snapshot) => {
      observed.push(snapshot?.revision);
      expect(store.getSnapshot()).toBe(snapshot);
    });

    store.clear();

    expect(observed).toEqual([undefined]);
  });
});
