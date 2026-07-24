import { describe, expect, it } from 'vitest';
import { buildSlicerValueIndex, compileSlicerFilterContext } from '../../../src/analysis/slicer';

describe('slicer filter context', () => {
  it('combines multi-select within a slicer as OR and slicers as AND per target', () => {
    const context = compileSlicerFilterContext([
      {
        id: 'region',
        fieldId: 'region-column',
        targets: ['table-1', 'chart-1'],
        selection: ['East', 'West'],
        stateScope: 'document',
      },
      {
        id: 'year',
        fieldId: 'year-column',
        targets: ['table-1'],
        selection: [2026],
        stateScope: 'document',
      },
    ]);

    expect(context.matches('table-1', { 'region-column': 'East', 'year-column': 2026 })).toBe(true);
    expect(context.matches('table-1', { 'region-column': 'North', 'year-column': 2026 })).toBe(
      false,
    );
    expect(context.matches('table-1', { 'region-column': 'West', 'year-column': 2025 })).toBe(
      false,
    );
    expect(context.matches('chart-1', { 'region-column': 'West' })).toBe(true);
  });

  it('uses session selection without mutating persistent slicer state', () => {
    const slicer = {
      id: 'region',
      fieldId: 'region-column',
      targets: ['table-1'],
      selection: ['East'],
      stateScope: 'session' as const,
    };
    const context = compileSlicerFilterContext([slicer], {
      region: ['West'],
    });

    expect(context.matches('table-1', { 'region-column': 'West' })).toBe(true);
    expect(context.matches('table-1', { 'region-column': 'East' })).toBe(false);
    expect(slicer.selection).toEqual(['East']);
  });

  it('treats an empty selection as no contribution and reports missing targets', () => {
    const context = compileSlicerFilterContext(
      [
        {
          id: 'empty',
          fieldId: 'region-column',
          targets: ['missing-table'],
          selection: [],
          stateScope: 'document',
        },
      ],
      {},
      { knownTargets: ['table-1'] },
    );

    expect(context.matches('table-1', {})).toBe(true);
    expect(context.diagnostics).toEqual([
      {
        code: 'SLICER_TARGET_MISSING',
        slicerId: 'empty',
        targetId: 'missing-table',
        message: 'Slicer empty references missing target missing-table',
      },
    ]);
  });

  it('indexes unique values deterministically and enforces cardinality limits', () => {
    expect(
      buildSlicerValueIndex(['West', 'East', 'West', null, 2, 1], { maximumValues: 8 }),
    ).toEqual([
      { value: null, count: 1 },
      { value: 1, count: 1 },
      { value: 2, count: 1 },
      { value: 'East', count: 1 },
      { value: 'West', count: 2 },
    ]);
    expect(() => buildSlicerValueIndex(['A', 'B'], { maximumValues: 1 })).toThrow(/value limit/u);
  });
});
