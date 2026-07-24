import { describe, expect, it } from 'vitest';
import { refreshPivot } from '../../../src/analysis/pivot';

const source = {
  revision: 'source-1',
  fields: ['region', 'quarter', 'amount', 'units'],
  rows: [
    ['West', 'Q2', 20, 2],
    ['East', 'Q1', 10, 1],
    ['West', 'Q1', 30, 3],
    ['East', 'Q1', 15, 2],
  ],
} as const;

const definition = {
  id: 'pivot-1',
  rows: ['region'],
  columns: ['quarter'],
  values: [
    { id: 'sales', field: 'amount', aggregate: 'sum' as const },
    { id: 'orders', field: 'amount', aggregate: 'count' as const },
    { id: 'average-units', field: 'units', aggregate: 'average' as const },
  ],
  filters: [],
};

describe('pivot refresh', () => {
  it('builds a deterministic immutable snapshot with lineage', async () => {
    const outcome = await refreshPivot(source, definition, {
      signal: new AbortController().signal,
    });

    expect(outcome.status).toBe('ready');
    if (outcome.status !== 'ready') return;
    expect(outcome.result.rowKeys).toEqual([['East'], ['West']]);
    expect(outcome.result.columnKeys).toEqual([['Q1'], ['Q2']]);
    expect(outcome.result.cells).toEqual([
      {
        rowKey: ['East'],
        columnKey: ['Q1'],
        values: { sales: 25, orders: 2, 'average-units': 1.5 },
        sourceRows: [1, 3],
      },
      {
        rowKey: ['West'],
        columnKey: ['Q1'],
        values: { sales: 30, orders: 1, 'average-units': 3 },
        sourceRows: [2],
      },
      {
        rowKey: ['West'],
        columnKey: ['Q2'],
        values: { sales: 20, orders: 1, 'average-units': 2 },
        sourceRows: [0],
      },
    ]);
    expect(Object.isFrozen(outcome.result.cells[0]?.values)).toBe(true);
  });

  it('applies fixed filter semantics before grouping', async () => {
    const outcome = await refreshPivot(
      source,
      {
        ...definition,
        filters: [{ field: 'quarter', values: ['Q1'] }],
      },
      { signal: new AbortController().signal },
    );

    expect(outcome.status === 'ready' && outcome.result.columnKeys).toEqual([['Q1']]);
  });

  it('preserves the previous successful result when cancelled', async () => {
    const initial = await refreshPivot(source, definition, {
      signal: new AbortController().signal,
    });
    if (initial.status !== 'ready') throw new Error('expected initial pivot result');
    const controller = new AbortController();
    controller.abort();

    const cancelled = await refreshPivot(source, definition, {
      signal: controller.signal,
      previous: initial.result,
    });

    expect(cancelled).toEqual({
      status: 'cancelled',
      stale: true,
      result: initial.result,
    });
  });

  it('does not retain a previous result from another definition', async () => {
    const initial = await refreshPivot(source, definition, {
      signal: new AbortController().signal,
    });
    if (initial.status !== 'ready') throw new Error('expected initial pivot result');
    const controller = new AbortController();
    controller.abort();

    await expect(
      refreshPivot(
        source,
        { ...definition, id: 'pivot-2' },
        { signal: controller.signal, previous: initial.result },
      ),
    ).resolves.toEqual({
      status: 'cancelled',
      stale: true,
    });
  });

  it('rejects unknown aggregates at the runtime boundary', async () => {
    await expect(
      refreshPivot(
        source,
        {
          ...definition,
          values: [
            {
              id: 'sales',
              field: 'amount',
              aggregate: 'median',
            },
          ],
        } as unknown as typeof definition,
        { signal: new AbortController().signal },
      ),
    ).rejects.toThrow(/aggregate/u);
  });

  it('bounds and can cancel deterministic post-processing', async () => {
    const manyGroups = {
      revision: 'source-many',
      fields: ['group', 'amount'],
      rows: Array.from({ length: 20 }, (_, index) => [`group-${index}`, index] as const),
    };
    const manyDefinition = {
      id: 'pivot-many',
      rows: ['group'],
      columns: [],
      values: [{ id: 'total', field: 'amount', aggregate: 'sum' as const }],
      filters: [],
    };

    await expect(
      refreshPivot(manyGroups, manyDefinition, {
        signal: new AbortController().signal,
        limits: { maximumPostProcessingSteps: 5 },
      }),
    ).rejects.toThrow(/post-processing/u);

    const controller = new AbortController();
    const outcome = await refreshPivot(manyGroups, manyDefinition, {
      signal: controller.signal,
      onProgress(completed, total) {
        if (completed === total) controller.abort();
      },
    });
    expect(outcome).toEqual({ status: 'cancelled', stale: true });
  });

  it('fails closed when result or source budgets are exceeded', async () => {
    await expect(
      refreshPivot(source, definition, {
        signal: new AbortController().signal,
        limits: { maximumRows: 2 },
      }),
    ).rejects.toThrow(/row limit/u);
    await expect(
      refreshPivot(source, definition, {
        signal: new AbortController().signal,
        limits: { maximumResultCells: 1 },
      }),
    ).rejects.toThrow(/result cell limit/u);
  });
});
