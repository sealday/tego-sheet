import { describe, expect, it } from 'vitest';
import {
  chartAffectedByChanges,
  chartToDisplayCommands,
  resolveChart,
  type ChartDefinition,
  type ChartValueSource,
} from '../../../src/analysis/charts';
import type { DocumentSheetId } from '../../../src/document';
import type { PrintDisplayCommand } from '../../../src/print';

const sheetId = 'sheet-1' as DocumentSheetId;
const otherSheetId = 'sheet-2' as DocumentSheetId;
const definition: ChartDefinition = {
  id: 'sales-chart',
  type: 'column',
  title: 'Quarterly sales',
  categories: {
    sheetId,
    start: { row: 1, column: 0 },
    end: { row: 3, column: 0 },
  },
  series: [
    {
      id: 'actual',
      name: 'Actual',
      values: {
        sheetId,
        start: { row: 1, column: 1 },
        end: { row: 3, column: 1 },
      },
    },
  ],
};

function source(values: Readonly<Record<string, readonly unknown[]>>): ChartValueSource {
  return {
    revision: 'revision-7',
    read: (reference) =>
      values[
        `${reference.sheetId}:${reference.start.row},${reference.start.column}-${reference.end.row},${reference.end.column}`
      ] ?? [],
  };
}

function flatten(commands: readonly PrintDisplayCommand[]): readonly PrintDisplayCommand[] {
  return commands.flatMap((command) =>
    command.kind === 'clip' || command.kind === 'group'
      ? [command, ...flatten(command.commands)]
      : [command],
  );
}

describe('normalized charts', () => {
  it('resolves stable range references into an immutable renderer-neutral model', () => {
    const result = resolveChart(
      definition,
      source({
        'sheet-1:1,0-3,0': ['Q1', 'Q2', 'Q3'],
        'sheet-1:1,1-3,1': [12, 18, 15],
      }),
    );

    expect(result.model).toEqual({
      id: 'sales-chart',
      type: 'column',
      title: 'Quarterly sales',
      sourceRevision: 'revision-7',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ id: 'actual', name: 'Actual', values: [12, 18, 15] }],
    });
    expect(result.dependencies).toEqual([definition.categories, definition.series[0]?.values]);
    expect(result.diagnostics).toEqual([]);
    expect(Object.isFrozen(result.model.series[0]?.values)).toBe(true);
  });

  it('emits deterministic diagnostics for invalid values and hard budgets', () => {
    const result = resolveChart(
      {
        ...definition,
        series: [
          definition.series[0]!,
          {
            id: 'forecast',
            values: {
              sheetId,
              start: { row: 1, column: 2 },
              end: { row: 3, column: 2 },
            },
          },
        ],
      },
      source({
        'sheet-1:1,0-3,0': ['Q1', 'Q2', 'Q3'],
        'sheet-1:1,1-3,1': [12, 'bad', 15],
        'sheet-1:1,2-3,2': [20, 22, 24],
      }),
      { maximumSeries: 1, maximumPoints: 3 },
    );

    expect(result.model.series).toEqual([{ id: 'actual', name: 'Actual', values: [12, null, 15] }]);
    expect(result.diagnostics.map(({ code }) => code)).toEqual([
      'CHART_SERIES_LIMIT_EXCEEDED',
      'CHART_NON_NUMERIC_VALUE',
    ]);
  });

  it('treats maximumPoints as a chart-wide budget and skips reads that cannot fit', () => {
    let forecastReads = 0;
    const result = resolveChart(
      {
        ...definition,
        categories: undefined,
        series: [
          definition.series[0]!,
          {
            id: 'forecast',
            values: {
              sheetId,
              start: { row: 1, column: 2 },
              end: { row: 3, column: 2 },
            },
          },
        ],
      },
      {
        revision: 'revision-8',
        read: (reference) => {
          if (reference.start.column === 2) forecastReads += 1;
          return [1, 2, 3];
        },
      },
      { maximumSeries: 2, maximumPoints: 4 },
    );

    expect(forecastReads).toBe(0);
    expect(result.model.series.map(({ values }) => values)).toEqual([[1, 2, 3], []]);
    expect(result.diagnostics.map(({ code }) => code)).toEqual(['CHART_POINT_LIMIT_EXCEEDED']);
  });

  it('invalidates only charts whose stable references intersect a change', () => {
    expect(
      chartAffectedByChanges(definition, [
        { sheetId, start: { row: 2, column: 1 }, end: { row: 2, column: 1 } },
      ]),
    ).toBe(true);
    expect(
      chartAffectedByChanges(definition, [
        { sheetId, start: { row: 20, column: 1 }, end: { row: 20, column: 1 } },
      ]),
    ).toBe(false);
    expect(
      chartAffectedByChanges(definition, [
        { sheetId: otherSheetId, start: { row: 2, column: 1 }, end: { row: 2, column: 1 } },
      ]),
    ).toBe(false);
  });

  it.each(['column', 'bar', 'line', 'pie'] as const)(
    'lays out %s charts as renderer-neutral vector commands with an accessible summary',
    (type) => {
      const result = resolveChart(
        { ...definition, type },
        source({
          'sheet-1:1,0-3,0': ['Q1', 'Q2', 'Q3'],
          'sheet-1:1,1-3,1': [12, 18, 15],
        }),
      );
      const output = chartToDisplayCommands(result.model, {
        x: 10,
        y: 20,
        width: 320,
        height: 180,
      });

      const commands = flatten(output.commands);
      expect(commands.length).toBeGreaterThan(2);
      expect(commands.every((command) => command.kind !== 'image')).toBe(true);
      expect(output.summary).toContain('Quarterly sales');
      expect(output.summary).toContain('Actual');
      expect(output.summary).toContain('Q1: 12');
    },
  );

  it('renders an isolated line point as a visible vector marker', () => {
    const result = resolveChart(
      {
        ...definition,
        type: 'line',
        categories: undefined,
        series: [
          {
            ...definition.series[0]!,
            values: { sheetId, start: { row: 1, column: 1 }, end: { row: 1, column: 1 } },
          },
        ],
      },
      source({ 'sheet-1:1,1-1,1': [12] }),
    );

    expect(
      flatten(
        chartToDisplayCommands(result.model, { x: 0, y: 0, width: 100, height: 60 }).commands,
      ).some((command) => command.kind === 'fill-rect'),
    ).toBe(true);
  });

  it('renders a one-slice pie with two arcs instead of a degenerate full-circle arc', () => {
    const result = resolveChart(
      {
        ...definition,
        type: 'pie',
        categories: undefined,
        series: [
          {
            ...definition.series[0]!,
            values: { sheetId, start: { row: 1, column: 1 }, end: { row: 1, column: 1 } },
          },
        ],
      },
      source({ 'sheet-1:1,1-1,1': [12] }),
    );
    const path = flatten(
      chartToDisplayCommands(result.model, { x: 0, y: 0, width: 100, height: 60 }).commands,
    ).find((command) => command.kind === 'path');

    expect(path?.kind === 'path' ? path.data.match(/\bA\b/gu)?.length : 0).toBe(2);
  });
});
