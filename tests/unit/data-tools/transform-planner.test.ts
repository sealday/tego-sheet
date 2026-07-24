import { describe, expect, it } from 'vitest';
import { createDataTransformPlanner } from '../../../src/data-tools';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  type DocumentSheetId,
} from '../../../src/document';
import type { SheetId } from '../../../src/core';
import { createDocumentController } from '../../../src/document-controller';

describe('DATA-01 revision-bound transform planner', () => {
  const documentSheetId = 'sheet-1' as DocumentSheetId;
  const sheetId = 'sheet-1' as SheetId;

  it('previews without mutation and commits one undoable transaction', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'seed',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 0 },
        text: 'draft invoice',
      },
    });
    const planner = createDataTransformPlanner({ maxCells: 100, maxSamples: 10 });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'find-replace',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
      find: 'draft',
      replacement: 'final',
      match: 'literal',
    });
    expect(preview.sampleChanges).toHaveLength(1);
    expect(controller.getSnapshot().revision).toBe(preview.baseRevision);

    expect(planner.commit(controller, preview.planId)).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'string',
      value: 'final invoice',
    });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
  });

  it('rejects a stale plan without modifying the newer revision', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const planner = createDataTransformPlanner({ maxCells: 100, maxSamples: 10 });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'find-replace',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
      find: 'x',
      replacement: 'y',
      match: 'literal',
    });
    controller.execute({
      schemaVersion: 1,
      id: 'newer',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 1, column: 0 },
        text: 'newer',
      },
    });
    expect(planner.commit(controller, preview.planId)).toMatchObject({
      status: 'rejected',
      code: 'TRANSFORM_PLAN_STALE',
    });
    expect(controller.getSnapshot().revision).toBe(1);
  });

  it('reports formula matches without rewriting formula source', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'formula',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 0 },
        text: '=A2+1',
      },
    });
    const planner = createDataTransformPlanner({
      maxCells: 100,
      maxCommands: 100,
      maxSamples: 10,
    });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'find-replace',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
      find: 'A2',
      replacement: 'B2',
      match: 'literal',
    });

    expect(preview).toMatchObject({
      affectedRange: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
      estimatedCellCount: 0,
      warnings: [
        expect.objectContaining({
          code: 'FORMULA_TRANSFORM_SKIPPED',
          location: {
            cell: { sheetId: documentSheetId, row: 0, column: 0 },
          },
        }),
      ],
    });
    expect(planner.commit(controller, preview.planId)).toMatchObject({ status: 'noop' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'formula',
      source: '=A2+1',
    });
  });

  it('previews and atomically commits bounded text splitting with overwrite warnings', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'seed',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 0 },
        text: 'Ada,Lovelace',
      },
    });
    controller.execute({
      schemaVersion: 1,
      id: 'occupied',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 1 },
        text: 'occupied',
      },
    });
    const planner = createDataTransformPlanner({
      maxCells: 100,
      maxCommands: 100,
      maxSamples: 10,
    });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'split-text',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 0, column: 0 },
      },
      delimiter: ',',
      maximumColumns: 2,
    });

    expect(preview.estimatedCellCount).toBe(2);
    expect(preview.warnings).toEqual([
      expect.objectContaining({
        code: 'NONEMPTY_TARGET_OVERWRITE',
        location: { cell: { sheetId: documentSheetId, row: 0, column: 1 } },
      }),
    ]);
    expect(preview.sampleChanges).toEqual([
      { row: 0, column: 0, before: 'Ada,Lovelace', after: 'Ada' },
      { row: 0, column: 1, before: 'occupied', after: 'Lovelace' },
    ]);

    expect(planner.commit(controller, preview.planId)).toMatchObject({ status: 'committed' });
    const cells = controller.getSnapshot().document.workbook.sheets[0]?.cells ?? [];
    expect(cells.map(({ cell }) => cell.input)).toEqual([
      { type: 'string', value: 'Ada' },
      { type: 'string', value: 'Lovelace' },
    ]);
    expect(controller.undo()).toMatchObject({ status: 'committed' });
  });

  it('removes duplicate rows through descending structural commands and one undo', async () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'document-1',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: ['alpha', 'beta', 'alpha', 'alpha'].map((value, row) => ({
              row,
              column: 0,
              cell: { input: { type: 'string' as const, value } },
            })),
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
    if (!parsed.ok) throw new Error('dedupe fixture must parse');
    const controller = createDocumentController(parsed.document);
    const planner = createDataTransformPlanner({
      maxCells: 100,
      maxCommands: 100,
      maxSamples: 10,
    });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'remove-duplicates',
      range: {
        sheetId: documentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 3, column: 0 },
      },
      keyColumns: [0],
      keep: 'first',
    });

    expect(preview.estimatedCellCount).toBe(2);
    expect(preview.sampleChanges.map(({ row }) => row)).toEqual([3, 2]);
    expect(planner.commit(controller, preview.planId)).toMatchObject({ status: 'committed' });
    expect(
      controller
        .getSnapshot()
        .document.workbook.sheets[0]?.cells.map(({ row, cell }) => [row, cell.input]),
    ).toEqual([
      [0, { type: 'string', value: 'alpha' }],
      [1, { type: 'string', value: 'beta' }],
    ]);
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells).toHaveLength(4);
  });

  it('fails closed on cancellation, invalid patterns, and command limits', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'seed',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId, row: 0, column: 0 },
        text: 'alpha',
      },
    });
    const planner = createDataTransformPlanner({
      maxCells: 100,
      maxCommands: 0,
      maxSamples: 10,
    });
    const aborted = new AbortController();
    aborted.abort();
    await expect(
      planner.preview(
        controller.getSnapshot(),
        {
          type: 'find-replace',
          range: {
            sheetId: documentSheetId,
            start: { row: 0, column: 0 },
            end: { row: 0, column: 0 },
          },
          find: 'alpha',
          replacement: 'beta',
          match: 'literal',
        },
        { signal: aborted.signal },
      ),
    ).rejects.toMatchObject({ code: 'TRANSFORM_ABORTED' });
    await expect(
      planner.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range: {
          sheetId: documentSheetId,
          start: { row: 0, column: 0 },
          end: { row: 0, column: 0 },
        },
        find: '(',
        replacement: '',
        match: 'regex',
      }),
    ).rejects.toMatchObject({ code: 'REPLACE_PATTERN_INVALID' });
    await expect(
      planner.preview(controller.getSnapshot(), {
        type: 'find-replace',
        range: {
          sheetId: documentSheetId,
          start: { row: 0, column: 0 },
          end: { row: 0, column: 0 },
        },
        find: 'alpha',
        replacement: 'beta',
        match: 'literal',
      }),
    ).rejects.toMatchObject({ code: 'TRANSFORM_TOO_LARGE' });
    expect(controller.getSnapshot().revision).toBe(1);
  });
});
