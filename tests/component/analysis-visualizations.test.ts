import { describe, expect, it } from 'vitest';
import {
  chartToDisplayCommands,
  resolveChart,
  resolveSparkline,
  sparklineToDisplayCommands,
} from '../../src/analysis';
import type { DocumentSheetId } from '../../src/document';
import { createFontMetrics } from '../../src/presentation';
import { createPrintDisplayList, validatePrintDisplayCommands } from '../../src/print';

const sheetId = 'sheet-1' as DocumentSheetId;

describe('analysis visualization output integration', () => {
  it('composes chart and sparkline vectors into the shared output display list', () => {
    const chart = resolveChart(
      {
        id: 'revenue',
        type: 'line',
        title: 'Revenue',
        series: [
          {
            id: 'actual',
            values: {
              sheetId,
              start: { row: 0, column: 0 },
              end: { row: 0, column: 2 },
            },
          },
        ],
      },
      { revision: 'r1', read: () => [4, 9, 7] },
    );
    const sparkline = resolveSparkline(
      {
        id: 'revenue-trend',
        type: 'column',
        source: {
          sheetId,
          start: { row: 0, column: 0 },
          end: { row: 0, column: 2 },
        },
        target: { sheetId, row: 0, column: 3 },
      },
      { revision: 'r1', read: () => [4, 9, 7] },
    );
    const chartOutput = chartToDisplayCommands(chart.model, {
      x: 10,
      y: 10,
      width: 240,
      height: 120,
    });
    const sparklineOutput = sparklineToDisplayCommands(sparkline.model, {
      x: 260,
      y: 10,
      width: 80,
      height: 24,
    });
    const overlays = [...chartOutput.commands, ...sparklineOutput.commands];

    expect(validatePrintDisplayCommands(overlays)).toEqual([]);
    const displayList = createPrintDisplayList({
      pages: [{ width: 360, height: 180, cells: [], overlays }],
      fontMetrics: createFontMetrics({
        fonts: {},
        fallbackFont: 'sans-serif',
        fallback: { averageAdvance: 7, lineHeight: 12 },
      }),
    });
    expect(displayList.diagnostics).toEqual([]);
    expect(displayList.pages[0]?.commands).toEqual(overlays);
    expect(chartOutput.summary).toContain('Revenue');
    expect(sparklineOutput.summary).toContain('revenue-trend');
  });
});
