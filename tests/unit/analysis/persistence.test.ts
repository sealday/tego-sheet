import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type DocumentSheetId,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createSpreadsheetDocument } from '../../../src/document/create-document';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';
import {
  chartToDisplayCommands,
  resolveChart,
  resolveSparkline,
  sparklineToDisplayCommands,
  type ChartDefinition,
  type SparklineDefinition,
} from '../../../src/analysis';

const documentSheetId = 'sheet-1' as DocumentSheetId;
const chart: ChartDefinition = {
  id: 'chart-sales',
  type: 'area',
  title: 'Sales',
  categories: {
    sheetId: documentSheetId,
    start: { row: 1, column: 0 },
    end: { row: 2, column: 0 },
  },
  series: [
    {
      id: 'actual',
      values: {
        sheetId: documentSheetId,
        start: { row: 1, column: 1 },
        end: { row: 2, column: 1 },
      },
    },
  ],
  anchor: {
    type: 'one-cell',
    cell: { sheetId: documentSheetId, row: 4, column: 0 },
    offset: { x: 0, y: 0 },
    size: { width: 300, height: 180 },
  },
  templateRepeat: 'shared',
};
const sparkline: SparklineDefinition = {
  id: 'spark-sales',
  type: 'line',
  source: {
    sheetId: documentSheetId,
    start: { row: 1, column: 1 },
    end: { row: 2, column: 1 },
  },
  target: { sheetId: documentSheetId, row: 1, column: 2 },
};

describe('persistent analysis visualizations', () => {
  it('keeps legacy documents valid when analysis collections are omitted', () => {
    const input = structuredClone(
      createSpreadsheetDocument({ id: 'legacy-analysis', sheetId: 'sheet-1' }),
    );
    const sheet = input.workbook.sheets[0] as unknown as Record<string, unknown>;
    delete sheet.charts;
    delete sheet.sparklines;

    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(true);
    if (!parsed.ok) throw new Error('legacy document must parse');
    expect(parsed.document.workbook.sheets[0]).toMatchObject({ charts: [], sparklines: [] });
  });

  it('rejects dangling analysis sheet references atomically', () => {
    const input = structuredClone(
      createSpreadsheetDocument({ id: 'dangling-analysis', sheetId: 'sheet-1' }),
    ) as unknown as SpreadsheetDocumentInput;
    input.workbook.sheets[0]!.sparklines = [
      {
        ...sparkline,
        source: { ...sparkline.source, sheetId: 'missing-sheet' as DocumentSheetId },
      },
    ];

    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(false);
    if (parsed.ok) throw new Error('dangling analysis reference must fail');
    expect(parsed.diagnostics).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: 'DANGLING_REFERENCE' })]),
    );
  });

  it('rejects chart anchors that leave the owning sheet or cross two-cell markers', () => {
    const input = structuredClone(
      createSpreadsheetDocument({ id: 'invalid-chart-anchors', sheetId: 'sheet-1' }),
    ) as unknown as SpreadsheetDocumentInput;
    input.workbook.sheets.push({
      id: 'sheet-2' as DocumentSheetId,
      name: 'Other',
      cells: [],
      merges: [],
    });
    input.workbook.sheets[0]!.charts = [
      {
        ...chart,
        anchor: {
          type: 'two-cell',
          from: {
            sheetId: 'sheet-2' as DocumentSheetId,
            row: 4,
            column: 3,
            offset: { x: 10, y: 10 },
          },
          to: {
            sheetId: documentSheetId,
            row: 2,
            column: 1,
            offset: { x: 0, y: 0 },
          },
        },
      },
    ];

    const parsed = parseSpreadsheetDocument(input);
    expect(parsed.ok).toBe(false);
    if (parsed.ok) throw new Error('invalid chart anchor must fail');
    expect(parsed.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'DOCUMENT_SCHEMA_INVALID',
          details: { path: '$.workbook.sheets[0].charts[0].anchor' },
        }),
        expect.objectContaining({
          code: 'OBJECT_ANCHOR_INVALID',
          details: { path: '$.workbook.sheets[0].charts[0].anchor' },
        }),
      ]),
    );
  });

  it('persists, commands, undoes, and structurally transforms chart and sparkline references', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'analysis-document', sheetId: 'sheet-1' }),
    );

    expect(
      controller.dispatch({ type: 'set-chart', sheet: sheetId('sheet-1'), chart }, 'ref').status,
    ).toBe('committed');
    expect(
      controller.dispatch({ type: 'set-sparkline', sheet: sheetId('sheet-1'), sparkline }, 'ref')
        .status,
    ).toBe('committed');

    const roundTrip = parseSpreadsheetDocument(
      serializeSpreadsheetDocument(controller.getDocument()),
    );
    expect(roundTrip.ok).toBe(true);
    if (!roundTrip.ok) throw new Error('analysis document must round-trip');
    expect(roundTrip.document.workbook.sheets[0]?.charts).toEqual([chart]);
    expect(roundTrip.document.workbook.sheets[0]?.sparklines).toEqual([sparkline]);

    expect(
      controller.dispatch(
        { type: 'insert-row', sheet: sheetId('sheet-1'), index: 1, count: 1 },
        'ref',
      ).status,
    ).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.charts[0]?.series[0]?.values).toMatchObject(
      {
        start: { row: 2 },
        end: { row: 3 },
      },
    );
    expect(controller.getDocument().workbook.sheets[0]?.sparklines[0]).toMatchObject({
      source: { start: { row: 2 }, end: { row: 3 } },
      target: { row: 2 },
    });

    expect(controller.undo('ref').status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.sparklines[0]).toEqual(sparkline);
  });

  it('uses the same immutable resolved model for screen, template, print, and output projections', () => {
    const resolvedChart = resolveChart(chart, {
      revision: 'r1',
      read: (range) => (range.start.column === 0 ? ['Q1', 'Q2'] : [10, 20]),
    }).model;
    const resolvedSparkline = resolveSparkline(sparkline, {
      revision: 'r1',
      read: () => [10, 20],
    }).model;
    const rect = { x: 0, y: 0, width: 300, height: 180 };
    const chartDisplay = chartToDisplayCommands(resolvedChart, rect);
    const sparkDisplay = sparklineToDisplayCommands(resolvedSparkline, rect);

    expect(chartToDisplayCommands(resolvedChart, rect)).toEqual(chartDisplay);
    expect(sparklineToDisplayCommands(resolvedSparkline, rect)).toEqual(sparkDisplay);
    expect(JSON.parse(JSON.stringify(chartDisplay))).toEqual(chartDisplay);
    expect(JSON.parse(JSON.stringify(sparkDisplay))).toEqual(sparkDisplay);
  });
});
