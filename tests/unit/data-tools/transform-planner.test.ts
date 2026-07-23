import { describe, expect, it } from 'vitest';
import { createDataTransformPlanner } from '../../../src/data-tools';
import { createSpreadsheetDocument } from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';

describe('DATA-01 revision-bound transform planner', () => {
  it('previews without mutation and commits one undoable transaction', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute({
      schemaVersion: 1,
      id: 'seed',
      command: {
        type: 'set-cell-text',
        address: { sheet: 'sheet-1', row: 0, column: 0 },
        text: 'draft invoice',
      },
    });
    const planner = createDataTransformPlanner({ maxCells: 100, maxSamples: 10 });
    const preview = await planner.preview(controller.getSnapshot(), {
      type: 'find-replace',
      range: {
        sheetId: 'sheet-1',
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
        sheetId: 'sheet-1',
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
        address: { sheet: 'sheet-1', row: 1, column: 0 },
        text: 'newer',
      },
    });
    expect(planner.commit(controller, preview.planId)).toMatchObject({
      status: 'rejected',
      code: 'TRANSFORM_PLAN_STALE',
    });
    expect(controller.getSnapshot().revision).toBe(1);
  });
});
