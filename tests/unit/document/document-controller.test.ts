import { describe, expect, it, vi } from 'vitest';
import {
  createDocumentController,
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type DocumentCommandEnvelope,
  type SheetId,
} from '../../../src';

function createMultiSheetDocument() {
  const input = JSON.parse(
    serializeSpreadsheetDocument(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    ),
  ) as {
    workbook: {
      sheets: Array<{ id: string; name: string; cells: unknown[]; merges: unknown[] }>;
    };
  };
  input.workbook.sheets.push({ id: 'sheet-2', name: 'Second', cells: [], merges: [] });
  const parsed = parseSpreadsheetDocument(input as never);
  if (!parsed.ok) throw new TypeError('Expected test document to parse');
  return parsed.document;
}

function setText(id: string, sheet: string, row: number, text: string): DocumentCommandEnvelope {
  return {
    schemaVersion: 1,
    id,
    command: {
      type: 'set-cell-text',
      address: { sheet: sheet as SheetId, row, column: 0 },
      text,
    },
  };
}

describe('public document controller', () => {
  it('exposes only stable document commands, transactions, snapshots, and events', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const listener = vi.fn();
    controller.subscribe(listener);

    const outcome = controller.transact({
      schemaVersion: 1,
      id: 'transaction-1',
      baseRevision: 0,
      commands: [
        setText('command-1', 'sheet-1', 0, 'alpha'),
        setText('command-2', 'sheet-1', 1, 'beta'),
      ],
    });

    expect(outcome).toMatchObject({
      status: 'committed',
      revision: 1,
      transaction: {
        id: 'transaction-1',
        committedRevision: 1,
        diagnostics: [],
      },
      change: {
        kind: 'transaction',
        commandCount: 2,
        sheets: [{ sheetId: 'sheet-1', kinds: ['cell'] }],
      },
    });
    expect(listener).toHaveBeenCalledTimes(1);
    expect(Object.keys(listener.mock.calls[0]?.[0] as object).sort()).toEqual([
      'change',
      'document',
      'revision',
      'transaction',
    ]);
    expect(controller.getSnapshot()).not.toHaveProperty('projection');
    expect(controller).not.toHaveProperty('replace');
    expect(controller).not.toHaveProperty('checkpoint');
  });

  it('supports dry-run, permission rejection, execute, and atomic history through the facade', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const envelope = {
      schemaVersion: 1,
      id: 'transaction-1',
      baseRevision: 0,
      commands: [setText('command-1', 'sheet-1', 0, 'alpha')],
    } as const;

    expect(controller.dryRun(envelope)).toMatchObject({ status: 'ready', baseRevision: 0 });
    expect(controller.getSnapshot().revision).toBe(0);
    expect(
      controller.transact(envelope, {
        permissionGate: ({ snapshot, transaction }) =>
          snapshot.revision > 0 && transaction.id === 'transaction-1',
      }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_NOT_ALLOWED' });
    expect(controller.execute(setText('command-2', 'sheet-1', 0, 'accepted'))).toMatchObject({
      status: 'committed',
      revision: 1,
    });
    expect(controller.undo()).toMatchObject({ status: 'committed', revision: 2 });
    expect(controller.redo()).toMatchObject({ status: 'committed', revision: 3 });
  });

  it('never throws after commits when public subscribers fail', () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.subscribe(() => {
      throw new Error('public observer failed');
    });

    expect(controller.execute(setText('command-1', 'sheet-1', 0, 'first'))).toMatchObject({
      status: 'committed',
      notificationError: 'public observer failed',
    });
    expect(controller.undo()).toMatchObject({
      status: 'committed',
      notificationError: 'public observer failed',
    });
    expect(controller.redo()).toMatchObject({
      status: 'committed',
      notificationError: 'public observer failed',
    });
    expect(
      controller.transact({
        schemaVersion: 1,
        id: 'transaction-2',
        baseRevision: 3,
        commands: [setText('command-2', 'sheet-1', 1, 'second')],
      }),
    ).toMatchObject({
      status: 'committed',
      notificationError: 'public observer failed',
    });
  });

  it('preserves multi-sheet aggregate changes across public undo and redo events', () => {
    const controller = createDocumentController(createMultiSheetDocument());
    const listener = vi.fn();
    controller.subscribe(listener);
    controller.transact({
      schemaVersion: 1,
      id: 'transaction-1',
      baseRevision: 0,
      commands: [
        setText('command-1', 'sheet-1', 0, 'first'),
        setText('command-2', 'sheet-2', 0, 'second'),
      ],
    });
    listener.mockClear();

    expect(controller.undo()).toMatchObject({
      status: 'committed',
      change: {
        kind: 'history',
        sheets: [{ sheetId: 'sheet-1' }, { sheetId: 'sheet-2' }],
      },
    });
    expect(listener.mock.calls[0]?.[0].change.sheets).toMatchObject([
      { sheetId: 'sheet-1' },
      { sheetId: 'sheet-2' },
    ]);
    listener.mockClear();
    expect(controller.redo()).toMatchObject({
      status: 'committed',
      change: {
        kind: 'history',
        sheets: [{ sheetId: 'sheet-1' }, { sheetId: 'sheet-2' }],
      },
    });
    expect(listener.mock.calls[0]?.[0].change.sheets).toMatchObject([
      { sheetId: 'sheet-1' },
      { sheetId: 'sheet-2' },
    ]);
  });
});
