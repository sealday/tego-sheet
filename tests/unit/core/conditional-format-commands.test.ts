import { expect, it } from 'vitest';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';
import { createSpreadsheetDocument } from '../../../src/document';
import { remapWorkbookCommand } from '../../../src/react/control/controlled-reconciler';

it('creates, updates, removes, undoes and redoes conditional formats as typed commands', () => {
  const controller = new SpreadsheetDocumentController(
    createSpreadsheetDocument({ id: 'conditional-commands', sheetId: 'sheet-1' }),
  );
  const sheet = sheetId('sheet-1');
  const range = {
    sheetId: 'sheet-1' as never,
    start: { row: 0, column: 0 },
    end: { row: 2, column: 0 },
  };

  expect(
    controller.dispatch(
      {
        type: 'set-conditional-format',
        sheet,
        index: 0,
        format: {
          type: 'cell-is',
          range,
          operator: 'greaterThan',
          formula: '10',
          style: { bold: true },
        },
      },
      'ref',
    ).status,
  ).toBe('committed');
  expect(controller.getDocument().workbook.sheets[0]?.conditionalFormatting).toHaveLength(1);

  controller.dispatch(
    {
      type: 'set-conditional-format',
      sheet,
      index: 0,
      format: {
        type: 'color-scale',
        range,
        minimumColor: '#000000',
        maximumColor: '#ffffff',
      },
    },
    'ref',
  );
  expect(controller.getDocument().workbook.sheets[0]?.conditionalFormatting[0]?.type).toBe(
    'color-scale',
  );

  controller.dispatch({ type: 'remove-conditional-format', sheet, index: 0 }, 'ref');
  expect(controller.getDocument().workbook.sheets[0]?.conditionalFormatting).toEqual([]);
  controller.dispatch({ type: 'undo' }, 'ref');
  expect(controller.getDocument().workbook.sheets[0]?.conditionalFormatting[0]?.type).toBe(
    'color-scale',
  );
  controller.dispatch({ type: 'redo' }, 'ref');
  expect(controller.getDocument().workbook.sheets[0]?.conditionalFormatting).toEqual([]);
});

it('remaps conditional-format ownership and qualified ranges for controlled replay', () => {
  const source = sheetId('sheet-1');
  const target = sheetId('sheet-2');
  const command = remapWorkbookCommand(
    {
      type: 'set-conditional-format',
      sheet: source,
      index: 0,
      format: {
        type: 'color-scale',
        range: {
          sheetId: 'sheet-1' as never,
          start: { row: 0, column: 0 },
          end: { row: 2, column: 0 },
        },
        minimumColor: '#000000',
        maximumColor: '#ffffff',
      },
    },
    new Map([[source, target]]),
  );

  expect(command).toMatchObject({
    type: 'set-conditional-format',
    sheet: target,
    format: { range: { sheetId: 'sheet-2' } },
  });
});
