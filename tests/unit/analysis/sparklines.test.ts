import { describe, expect, it } from 'vitest';
import {
  resolveSparkline,
  sparklineAffectedByChanges,
  sparklineToDisplayCommands,
  type SparklineDefinition,
} from '../../../src/analysis/sparklines';
import type { DocumentSheetId } from '../../../src/document';

const sheetId = 'sheet-1' as DocumentSheetId;
const definition: SparklineDefinition = {
  id: 'trend-a',
  type: 'line',
  source: {
    sheetId,
    start: { row: 0, column: 0 },
    end: { row: 0, column: 3 },
  },
  target: { sheetId, row: 0, column: 4 },
};

describe('sparklines', () => {
  it('normalizes values without changing or representing the target cell value', () => {
    const result = resolveSparkline(definition, {
      revision: 'r2',
      read: () => [1, 3, null, -2],
    });

    expect(result.model).toEqual({
      id: 'trend-a',
      type: 'line',
      sourceRevision: 'r2',
      values: [1, 3, null, -2],
      target: definition.target,
    });
    expect('cellValue' in result.model).toBe(false);
    expect(result.diagnostics).toEqual([]);
  });

  it('enforces the point budget before reading an oversized source', () => {
    let reads = 0;
    const result = resolveSparkline(
      {
        ...definition,
        source: {
          sheetId,
          start: { row: 0, column: 0 },
          end: { row: 0, column: 100 },
        },
      },
      {
        revision: 'r3',
        read: () => {
          reads += 1;
          return [];
        },
      },
      { maximumPoints: 10 },
    );

    expect(reads).toBe(0);
    expect(result.model.values).toEqual([]);
    expect(result.diagnostics.map(({ code }) => code)).toEqual(['SPARKLINE_POINT_LIMIT_EXCEEDED']);
  });

  it.each(['line', 'column', 'win-loss'] as const)(
    'lays out %s sparklines inside the supplied cell rectangle',
    (type) => {
      const result = resolveSparkline(
        { ...definition, type },
        { revision: 'r4', read: () => [-2, 0, 3, 1] },
      );
      const output = sparklineToDisplayCommands(result.model, {
        x: 5,
        y: 6,
        width: 80,
        height: 20,
      });

      expect(output.commands.length).toBeGreaterThan(0);
      expect(output.commands.every((command) => command.kind !== 'text')).toBe(true);
      expect(output.summary).toContain('trend-a');
      expect(output.summary).toContain('-2, 0, 3, 1');
    },
  );

  it('tracks only its source dependency, never the target cell', () => {
    expect(
      sparklineAffectedByChanges(definition, [
        { sheetId, start: { row: 0, column: 2 }, end: { row: 0, column: 2 } },
      ]),
    ).toBe(true);
    expect(
      sparklineAffectedByChanges(definition, [
        { sheetId, start: { row: 0, column: 4 }, end: { row: 0, column: 4 } },
      ]),
    ).toBe(false);
  });

  it('renders isolated line points as visible markers without bridging blank gaps', () => {
    const single = resolveSparkline(definition, {
      revision: 'r5',
      read: () => [3],
    });
    expect(
      sparklineToDisplayCommands(single.model, { x: 0, y: 0, width: 40, height: 20 }).commands.some(
        (command) => command.kind === 'fill-rect',
      ),
    ).toBe(true);

    const gaps = resolveSparkline(definition, {
      revision: 'r6',
      read: () => [1, null, 3, 4],
    });
    const path = sparklineToDisplayCommands(gaps.model, {
      x: 0,
      y: 0,
      width: 40,
      height: 20,
    }).commands.find((command) => command.kind === 'path');
    expect(path?.kind === 'path' ? path.data.match(/\bM\b/gu)?.length : 0).toBe(2);
  });
});
