import { describe, expect, it, vi } from 'vitest';
import { createSpreadsheetDocument } from '../../../src/document';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';
import { sheetId } from '../../../src/core';

describe('SpreadsheetDocumentController', () => {
  it('rejects invalid schema 2 input instead of accepting legacy workbook data', () => {
    expect(
      () => new SpreadsheetDocumentController([{ name: 'legacy mutable workbook' }] as never),
    ).toThrow(/spreadsheet document/i);
  });

  it('isolates and deeply freezes input, snapshots, and subscriber documents', () => {
    const input = createSpreadsheetDocument({
      id: 'document-1',
      sheetId: 'sheet-1',
      sheetName: 'Input',
    });
    const controller = new SpreadsheetDocumentController(input);
    const received = vi.fn();
    controller.subscribe(received);

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'committed',
      },
      'ref',
    );

    const snapshot = controller.getDocument();
    const emitted = received.mock.lastCall?.[0].commit.document;
    expect(snapshot).not.toBe(input);
    expect(emitted).not.toBe(snapshot);
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.workbook.sheets[0]?.cells)).toBe(true);
    expect(Object.isFrozen(emitted)).toBe(true);
  });

  it('publishes exactly one document event per committed command', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const received = vi.fn();
    controller.subscribe(received);

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'one',
      },
      'ref',
    );

    expect(received).toHaveBeenCalledOnce();
    expect(received.mock.calls[0]?.[0].snapshot.revision).toBe(1);
  });

  it('runs the atomic before-notify observer exactly once', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const beforeNotify = vi.fn();

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'one',
      },
      'ref',
      { beforeNotify },
    );

    expect(beforeNotify).toHaveBeenCalledOnce();
  });

  it('preserves the schema 2 sheet ID through rename and edit commands', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({
        id: 'document-1',
        sheetId: 'stable-sheet',
        sheetName: 'Before',
      }),
    );
    const stableSheet = sheetId('stable-sheet');

    controller.dispatch({ type: 'rename-sheet', sheet: stableSheet, name: 'After' }, 'ref');
    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: stableSheet, row: 2, column: 3 },
        text: 'kept',
      },
      'ref',
    );

    const sheet = controller.getDocument().workbook.sheets[0];
    expect(sheet?.id).toBe('stable-sheet');
    expect(sheet?.name).toBe('After');
    expect(sheet?.cells[0]).toMatchObject({
      row: 2,
      column: 3,
      cell: { input: { type: 'string', value: 'kept' } },
    });
  });
});
