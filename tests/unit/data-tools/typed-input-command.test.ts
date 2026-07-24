import { describe, expect, it, vi } from 'vitest';
import type { SheetId } from '../../../src/core';
import { createSpreadsheetDocument } from '../../../src/document';
import {
  createDocumentController,
  type DocumentCommandEnvelope,
} from '../../../src/document-controller';

const sheet = 'sheet-1' as SheetId;

function typedCommand(id: string, input: unknown, row = 0): DocumentCommandEnvelope {
  return {
    schemaVersion: 1,
    id,
    command: {
      type: 'set-cell-input',
      address: { sheet, row, column: 0 },
      input,
    },
  } as unknown as DocumentCommandEnvelope;
}

describe('versioned typed cell input command', () => {
  it('commits a finite number atomically and preserves its type through undo and redo', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'typed-input-document', sheetId: 'sheet-1' }),
    );
    const permissionGate = vi.fn(() => true);
    const command = typedCommand('typed-number', { type: 'number', value: 42 });
    expect(JSON.parse(JSON.stringify(command))).toEqual(command);

    expect(
      controller.transact(
        {
          schemaVersion: 1,
          id: 'typed-input-transaction',
          baseRevision: 0,
          commands: [command],
        },
        { permissionGate },
      ),
    ).toMatchObject({
      status: 'committed',
      transaction: {
        commands: [
          {
            command: {
              type: 'set-cell-input',
              input: { type: 'number', value: 42 },
            },
          },
        ],
      },
    });
    expect(permissionGate).toHaveBeenCalledTimes(1);
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'number',
      value: 42,
    });

    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells).toEqual([]);
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'number',
      value: 42,
    });
  });

  it('supports normalized scalar/formula inputs and fails closed on invalid typed payloads', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'typed-input-document', sheetId: 'sheet-1' }),
    );
    for (const [row, input] of [
      [0, { type: 'blank' }],
      [1, { type: 'string', value: 'text' }],
      [2, { type: 'boolean', value: true }],
      [3, { type: 'formula', source: '=1+1' }],
    ] as const) {
      expect(controller.execute(typedCommand(`typed-${row}`, input, row))).toMatchObject({
        status: 'committed',
      });
    }
    expect(
      controller.getSnapshot().document.workbook.sheets[0]?.cells.map(({ cell }) => cell.input),
    ).toEqual([
      { type: 'blank' },
      { type: 'string', value: 'text' },
      { type: 'boolean', value: true },
      { type: 'formula', source: '=1+1' },
    ]);

    expect(
      controller.execute(typedCommand('bad-number', { type: 'number', value: Number.NaN }, 4)),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(
      controller.execute(typedCommand('bad-formula', { type: 'formula', source: '1+1' }, 4)),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(
      controller.execute(typedCommand('bad-shape', { type: 'number', value: '42' }, 4)),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(
      controller.execute(typedCommand('bad-extra-key', { type: 'blank', value: null }, 4)),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
  });

  it('commits a type-only change when the legacy display text is unchanged', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'typed-input-document', sheetId: 'sheet-1' }),
    );
    expect(
      controller.execute(typedCommand('string-number', { type: 'string', value: '42' })),
    ).toMatchObject({ status: 'committed' });

    expect(
      controller.execute(typedCommand('typed-number', { type: 'number', value: 42 })),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'number',
      value: 42,
    });
  });
});
