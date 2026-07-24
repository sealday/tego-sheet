import { describe, expect, it } from 'vitest';
import type { SheetId } from '../../../src/core';
import { createDataTransformPlanner, type AutofillTransform } from '../../../src/data-tools';
import {
  parseSpreadsheetDocument,
  type CellInput,
  type DocumentSheetId,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';

const documentSheet = 'sheet-1' as DocumentSheetId;
const sheet = 'sheet-1' as SheetId;

function controllerFixture() {
  const inputs: readonly (readonly CellInput[])[] = [
    [
      { type: 'number', value: 1 },
      { type: 'string', value: '2024-01-01' },
      { type: 'string', value: 'Item01' },
      { type: 'formula', source: '=$A1+B$1' },
      { type: 'boolean', value: true },
    ],
    [
      { type: 'number', value: 3 },
      { type: 'string', value: '2024-01-03' },
      { type: 'string', value: 'Item03' },
      { type: 'formula', source: '=$A2+B$1' },
      { type: 'boolean', value: true },
    ],
  ];
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'enhanced-autofill-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          rowCount: 6_000,
          columnCount: 5,
          cells: inputs.flatMap((row, rowIndex) =>
            row.map((input, column) => ({
              row: rowIndex,
              column,
              cell: { input, metadata: { seed: true } },
            })),
          ),
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
  } as SpreadsheetDocumentInput);
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return createDocumentController(parsed.document);
}

function request(endRow = 4): AutofillTransform {
  return {
    type: 'autofill',
    source: {
      sheetId: documentSheet,
      start: { row: 0, column: 0 },
      end: { row: 1, column: 4 },
    },
    target: {
      sheetId: documentSheet,
      start: { row: 2, column: 0 },
      end: { row: endRow, column: 4 },
    },
    mode: 'all',
  };
}

function inputAt(
  controller: ReturnType<typeof controllerFixture>,
  row: number,
  column: number,
): CellInput | undefined {
  return controller
    .getSnapshot()
    .document.workbook.sheets[0]?.cells.find(
      (entry) => entry.row === row && entry.column === column,
    )?.cell.input;
}

describe('DATA-01 enhanced autofill planner', () => {
  it('uses the command canonical expanded target for preview, budgets, samples, and commit', async () => {
    const controller = controllerFixture();
    const compact: AutofillTransform = {
      ...request(),
      target: {
        sheetId: documentSheet,
        start: { row: 2, column: 0 },
        end: { row: 2, column: 0 },
      },
    };
    const planner = createDataTransformPlanner({
      maxCells: 20,
      maxCommands: 1,
      maxSamples: 10,
    });
    const preview = await planner.preview(controller.getSnapshot(), compact);

    expect(preview).toMatchObject({
      affectedRange: {
        sheetId: documentSheet,
        start: { row: 2, column: 0 },
        end: { row: 3, column: 4 },
      },
      estimatedCellCount: 10,
    });
    expect(preview.sampleChanges).toHaveLength(10);
    expect(preview.sampleChanges.at(-1)).toEqual({
      row: 3,
      column: 4,
      before: '',
      after: 'true',
    });
    expect(planner.commit(controller, preview.planId)).toMatchObject({ status: 'committed' });
    expect(inputAt(controller, 3, 0)).toEqual({ type: 'number', value: 7 });
    expect(inputAt(controller, 3, 3)).toEqual({
      type: 'formula',
      source: '=$A4+B$1',
    });

    await expect(
      createDataTransformPlanner({ maxCells: 19, maxCommands: 1, maxSamples: 10 }).preview(
        controller.getSnapshot(),
        compact,
      ),
    ).rejects.toMatchObject({ code: 'TRANSFORM_TOO_LARGE' });
    await expect(
      createDataTransformPlanner({ maxCells: 20, maxCommands: 0, maxSamples: 10 }).preview(
        controller.getSnapshot(),
        compact,
      ),
    ).rejects.toMatchObject({ code: 'TRANSFORM_TOO_LARGE' });
  });

  it('previews and atomically commits typed patterns and AST-translated formulas', async () => {
    const controller = controllerFixture();
    const planner = createDataTransformPlanner({
      maxCells: 100,
      maxCommands: 10,
      maxSamples: 5,
    });
    const revision = controller.getSnapshot().revision;
    const preview = await planner.preview(controller.getSnapshot(), request());

    expect(controller.getSnapshot().revision).toBe(revision);
    expect(preview).toMatchObject({
      baseRevision: revision,
      affectedRange: request().target,
      estimatedCellCount: 15,
      sampleChanges: [
        { row: 2, column: 0, before: '', after: '5' },
        { row: 2, column: 1, before: '', after: '2024-01-05' },
        { row: 2, column: 2, before: '', after: 'Item05' },
        { row: 2, column: 3, before: '', after: '=$A3+B$1' },
        { row: 2, column: 4, before: '', after: 'true' },
      ],
    });

    expect(planner.commit(controller, preview.planId)).toMatchObject({
      status: 'committed',
      transaction: { commands: [{ command: { type: 'autofill' } }] },
    });
    expect(inputAt(controller, 2, 0)).toEqual({ type: 'number', value: 5 });
    expect(inputAt(controller, 4, 0)).toEqual({ type: 'number', value: 9 });
    expect(inputAt(controller, 2, 1)).toEqual({ type: 'string', value: '2024-01-05' });
    expect(inputAt(controller, 2, 2)).toEqual({ type: 'string', value: 'Item05' });
    expect(inputAt(controller, 2, 3)).toEqual({
      type: 'formula',
      source: '=$A3+B$1',
    });
    expect(inputAt(controller, 4, 3)).toEqual({
      type: 'formula',
      source: '=$A5+B$1',
    });
    expect(inputAt(controller, 2, 4)).toEqual({ type: 'boolean', value: true });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(inputAt(controller, 2, 0)).toBeUndefined();
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(inputAt(controller, 2, 0)).toEqual({ type: 'number', value: 5 });
  });

  it('rejects stale plans and yields so post-invocation cancellation publishes no plan', async () => {
    const controller = controllerFixture();
    const planner = createDataTransformPlanner({
      maxCells: 30_000,
      maxCommands: 10,
      maxSamples: 5,
    });
    const stale = await planner.preview(controller.getSnapshot(), request());
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'advance-revision',
        command: {
          type: 'set-cell-input',
          address: { sheet, row: 5_500, column: 0 },
          input: { type: 'number', value: 1 },
        },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(planner.commit(controller, stale.planId)).toEqual({
      status: 'rejected',
      code: 'TRANSFORM_PLAN_STALE',
    });

    const abortController = new AbortController();
    const pending = planner.preview(controller.getSnapshot(), request(4_999), {
      signal: abortController.signal,
    });
    queueMicrotask(() => abortController.abort());
    await expect(pending).rejects.toMatchObject({ code: 'TRANSFORM_ABORTED' });
  });

  it('fails before publication when the combined source and target exceed the cell budget', async () => {
    const controller = controllerFixture();
    const planner = createDataTransformPlanner({ maxCells: 24, maxCommands: 10, maxSamples: 5 });
    await expect(planner.preview(controller.getSnapshot(), request())).rejects.toMatchObject({
      code: 'TRANSFORM_TOO_LARGE',
    });
  });
});
