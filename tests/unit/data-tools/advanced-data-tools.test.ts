import { describe, expect, it, vi } from 'vitest';
import type { SheetId } from '../../../src/core';
import {
  analyzeDataAnomalies,
  createDataTransformPlanner,
  type DataTransformPreview,
} from '../../../src/data-tools';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  type DocumentSheetId,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';

const documentSheetId = 'sheet-1' as DocumentSheetId;
const sheetId = 'sheet-1' as SheetId;
const range = {
  sheetId: documentSheetId,
  start: { row: 0, column: 0 },
  end: { row: 0, column: 0 },
} as const;

function seededController(
  cells: readonly {
    readonly row: number;
    readonly column: number;
    readonly input:
      | { readonly type: 'string'; readonly value: string }
      | { readonly type: 'number'; readonly value: number }
      | { readonly type: 'formula'; readonly source: string };
  }[],
) {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'data-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: cells.map(({ row, column, input }) => ({ row, column, cell: { input } })),
          merges: [],
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error('data fixture must parse');
  return createDocumentController(parsed.document);
}

describe('DATA-01 advanced cleanup safety', () => {
  it.each(['(a+)+$', '(?=a)a', String.raw`(a)\1`])(
    'rejects unsafe regular expression %s before evaluation',
    async (find) => {
      const controller = seededController([
        { row: 0, column: 0, input: { type: 'string', value: 'a' } },
      ]);
      const planner = createDataTransformPlanner({
        maxCells: 10,
        maxSamples: 10,
        maxRegexInputLength: 100,
        maxRegexSteps: 1_000,
        maxRegexMilliseconds: 10,
      });

      await expect(
        planner.preview(controller.getSnapshot(), {
          type: 'find-replace',
          range,
          find,
          replacement: 'x',
          match: 'regex',
        }),
      ).rejects.toMatchObject({ code: 'REPLACE_PATTERN_INVALID' });
    },
  );

  it('enforces regex input and estimated-step budgets while allowing conservative safe patterns', async () => {
    const controller = seededController([
      { row: 0, column: 0, input: { type: 'string', value: 'aaaaab' } },
    ]);
    const constrained = createDataTransformPlanner({
      maxCells: 10,
      maxSamples: 10,
      maxRegexInputLength: 5,
      maxRegexSteps: 1_000,
      maxRegexMilliseconds: 10,
    });
    await expect(
      constrained.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range,
        find: 'a+b',
        replacement: 'ok',
        match: 'regex',
      }),
    ).rejects.toMatchObject({ code: 'REPLACE_BUDGET_EXCEEDED' });

    const stepBounded = createDataTransformPlanner({
      maxCells: 10,
      maxSamples: 10,
      maxRegexPatternLength: 100,
      maxRegexInputLength: 100,
      maxRegexSteps: 5,
      maxRegexMilliseconds: 10,
    });
    await expect(
      stepBounded.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range,
        find: 'a+b',
        replacement: 'ok',
        match: 'regex',
      }),
    ).rejects.toMatchObject({ code: 'REPLACE_BUDGET_EXCEEDED' });

    const patternBounded = createDataTransformPlanner({
      maxCells: 10,
      maxSamples: 10,
      maxRegexPatternLength: 2,
      maxRegexInputLength: 100,
      maxRegexSteps: 1_000,
      maxRegexMilliseconds: 10,
    });
    await expect(
      patternBounded.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range,
        find: 'a+b',
        replacement: 'ok',
        match: 'regex',
      }),
    ).rejects.toMatchObject({ code: 'REPLACE_PATTERN_INVALID' });

    const safe = createDataTransformPlanner({
      maxCells: 10,
      maxSamples: 10,
      maxRegexInputLength: 100,
      maxRegexSteps: 1_000,
      maxRegexMilliseconds: 10,
    });
    await expect(
      safe.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range,
        find: 'a+b',
        replacement: 'ok',
        match: 'regex',
      }),
    ).resolves.toMatchObject({
      sampleChanges: [{ before: 'aaaaab', after: 'ok' }],
    });
  });

  it('fails closed when the cumulative regex time budget is exceeded', async () => {
    const controller = seededController([
      { row: 0, column: 0, input: { type: 'string', value: 'aaaaab' } },
    ]);
    const clock = vi.spyOn(performance, 'now');
    clock.mockReturnValueOnce(0).mockReturnValueOnce(20);
    try {
      const planner = createDataTransformPlanner({
        maxCells: 10,
        maxSamples: 10,
        maxRegexInputLength: 100,
        maxRegexSteps: 1_000,
        maxRegexMilliseconds: 10,
      });
      await expect(
        planner.preview(controller.getSnapshot(), {
          type: 'find-replace',
          range,
          find: 'a+b',
          replacement: 'ok',
          match: 'regex',
        }),
      ).rejects.toMatchObject({ code: 'REPLACE_BUDGET_EXCEEDED' });
    } finally {
      clock.mockRestore();
    }
  });

  it('fills bounded numeric, date, and text-suffix sequences through one undoable transaction', async () => {
    const controller = seededController([
      { row: 0, column: 0, input: { type: 'string', value: 'wrong type' } },
    ]);
    const planner = createDataTransformPlanner({ maxCells: 20, maxCommands: 20, maxSamples: 20 });
    const numeric = await planner.preview(controller.getSnapshot(), {
      type: 'fill-series',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 3, column: 0 },
      },
      series: 'number',
      seed: ['1', '3'],
    });

    expect(numeric.sampleChanges.map(({ after }) => after)).toEqual(['1', '3', '5', '7']);
    expect(numeric.warnings).toEqual([
      expect.objectContaining({
        code: 'FILL_SERIES_TYPE_OVERWRITE',
        location: { cell: { sheetId: documentSheetId, row: 0, column: 0 } },
      }),
    ]);
    expect(planner.commit(controller, numeric.planId)).toMatchObject({ status: 'committed' });
    expect(
      controller.getSnapshot().document.workbook.sheets[0]?.cells.map(({ cell }) => cell.input),
    ).toEqual([
      { type: 'string', value: '1' },
      { type: 'string', value: '3' },
      { type: 'string', value: '5' },
      { type: 'string', value: '7' },
    ]);
    expect(controller.undo()).toMatchObject({ status: 'committed' });

    const date = await planner.preview(controller.getSnapshot(), {
      type: 'fill-series',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 1 },
        end: { row: 0, column: 3 },
      },
      series: 'date',
      seed: ['2024-01-01', '2024-01-03'],
    });
    expect(date.sampleChanges.map(({ after }) => after)).toEqual([
      '2024-01-01',
      '2024-01-03',
      '2024-01-05',
    ]);

    const suffix = await planner.preview(controller.getSnapshot(), {
      type: 'fill-series',
      range: {
        sheetId: documentSheetId,
        start: { row: 1, column: 1 },
        end: { row: 1, column: 3 },
      },
      series: 'text-suffix',
      seed: ['=ITEM01', '=ITEM03'],
    });
    expect(suffix.sampleChanges.map(({ after }) => after)).toEqual([
      '=ITEM01',
      '=ITEM03',
      '=ITEM05',
    ]);
    expect(suffix.warnings).toEqual([
      expect.objectContaining({ code: 'FORMULA_INJECTION_RISK' }),
      expect.objectContaining({ code: 'FORMULA_INJECTION_RISK' }),
      expect.objectContaining({ code: 'FORMULA_INJECTION_RISK' }),
    ]);
  });

  it('returns immutable read-only blank, error, and type-outlier findings with abort and size limits', async () => {
    const controller = seededController([
      { row: 0, column: 0, input: { type: 'string', value: 'expected' } },
      { row: 1, column: 0, input: { type: 'number', value: 7 } },
      { row: 2, column: 0, input: { type: 'formula', source: '=#REF!' } },
    ]);
    const snapshot = controller.getSnapshot();
    const result = await analyzeDataAnomalies(
      snapshot,
      {
        range: {
          sheetId: documentSheetId,
          start: { row: 0, column: 0 },
          end: { row: 3, column: 0 },
        },
        checks: ['blank', 'error', 'type-outlier'],
        expectedType: 'string',
      },
      {
        maxCells: 10,
        maxFindings: 10,
        context: {
          errorCells: [{ sheetId: documentSheetId, row: 2, column: 0 }],
        },
      },
    );

    expect(result).toEqual({
      inspectedCellCount: 4,
      truncated: false,
      findings: [
        expect.objectContaining({
          code: 'DATA_TYPE_OUTLIER',
          location: { cell: { sheetId: documentSheetId, row: 1, column: 0 } },
        }),
        expect.objectContaining({
          code: 'DATA_ERROR_ANOMALY',
          location: { cell: { sheetId: documentSheetId, row: 2, column: 0 } },
        }),
        expect.objectContaining({
          code: 'DATA_BLANK_ANOMALY',
          location: { cell: { sheetId: documentSheetId, row: 3, column: 0 } },
        }),
      ],
    });
    expect('planId' in result).toBe(false);
    expect(Object.isFrozen(result.findings)).toBe(true);
    expect(Object.isFrozen(result.findings[0])).toBe(true);
    expect(controller.getSnapshot().revision).toBe(snapshot.revision);

    const aborted = new AbortController();
    aborted.abort();
    await expect(
      analyzeDataAnomalies(
        snapshot,
        { range, checks: ['blank'] },
        { maxCells: 10, maxFindings: 10, signal: aborted.signal },
      ),
    ).rejects.toMatchObject({ code: 'TRANSFORM_ABORTED' });
    await expect(
      analyzeDataAnomalies(
        snapshot,
        {
          range: {
            sheetId: documentSheetId,
            start: { row: 0, column: 0 },
            end: { row: 10, column: 0 },
          },
          checks: ['blank'],
        },
        { maxCells: 10, maxFindings: 10 },
      ),
    ).rejects.toMatchObject({ code: 'TRANSFORM_TOO_LARGE' });
  });

  it('warns through optional preview context when changes intersect template regions', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'seed',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 0 },
        text: 'draft',
      },
    });
    const planner = createDataTransformPlanner({ maxCells: 10, maxSamples: 10 });
    const preview: DataTransformPreview = await planner.preview(
      controller.getSnapshot(),
      {
        type: 'find-replace',
        range,
        find: 'draft',
        replacement: 'final',
        match: 'literal',
      },
      { context: { templateRegions: [range] } },
    );

    expect(preview.warnings).toContainEqual(
      expect.objectContaining({
        code: 'TEMPLATE_REGION_CONFLICT',
        location: { range },
      }),
    );
  });
});
