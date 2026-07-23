import { describe, expect, it, vi } from 'vitest';
import { createSpreadsheetDocument, serializeSpreadsheetDocument } from '../../../src/document';
import { sheetId } from '../../../src/core';
import {
  SpreadsheetDocumentController,
  type SerializableCommandEnvelope,
  type SerializableTransactionEnvelope,
} from '../../../src/core/controller/spreadsheet-document-controller';

function command(id: string, row: number, text: string): SerializableCommandEnvelope {
  return {
    schemaVersion: 1,
    id,
    command: {
      type: 'set-cell-text',
      address: { sheet: sheetId('sheet-1'), row, column: 0 },
      text,
    },
  };
}

function transaction(
  controller: SpreadsheetDocumentController,
  commands: readonly SerializableCommandEnvelope[],
): SerializableTransactionEnvelope {
  return {
    schemaVersion: 1,
    id: 'transaction-1',
    baseRevision: controller.getSnapshot().revision,
    commands,
    metadata: { actor: 'unit-test' },
  };
}

describe('SpreadsheetDocumentController transactions', () => {
  it('commits multiple commands as one revision, event, and undo item', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const listener = vi.fn();
    controller.subscribe(listener);

    const outcome = controller.transact(
      transaction(controller, [command('command-1', 0, 'alpha'), command('command-2', 1, 'beta')]),
      { source: 'ref' },
    );

    expect(outcome.status).toBe('committed');
    if (outcome.status === 'committed') expect(outcome.change.kind).toBe('transaction');
    expect(controller.getSnapshot().revision).toBe(1);
    expect(controller.historySize).toEqual({ undo: 1, redo: 0 });
    expect(listener).toHaveBeenCalledTimes(1);
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 0, column: 0 })).toBe('alpha');
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 1, column: 0 })).toBe('beta');
  });

  it('rejects revision conflicts and permission denials before mutation', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = controller.getSnapshot();
    const stale = {
      ...transaction(controller, [command('command-1', 0, 'stale')]),
      baseRevision: 99,
    };

    expect(controller.transact(stale)).toMatchObject({
      status: 'rejected',
      code: 'REVISION_CONFLICT',
    });
    expect(
      controller.transact(transaction(controller, [command('command-2', 0, 'denied')]), {
        permissionGate: () => false,
      }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_NOT_ALLOWED' });
    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
  });

  it('prevents permission gates from reentering the mutation boundary', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const outcome = controller.transact(transaction(controller, [command('outer', 0, 'outer')]), {
      permissionGate: () => {
        controller.dispatch(command('inner', 1, 'inner').command, 'ref');
        return true;
      },
    });

    expect(outcome).toMatchObject({ status: 'rejected', code: 'COMMAND_NOT_ALLOWED' });
    expect(controller.getSnapshot().revision).toBe(0);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
  });

  it('isolates and validates envelopes, IDs, limits, JSON metadata, and gate failures', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const duplicate = transaction(controller, [
      command('duplicate', 0, 'first'),
      command('duplicate', 1, 'second'),
    ]);
    expect(controller.transact(duplicate)).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
    });
    expect(
      controller.transact({
        ...transaction(controller, [command('command-1', 0, 'metadata')]),
        metadata: { callback: (() => undefined) as never },
      }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(
      controller.transact({
        ...transaction(controller, [command('command-1', 0, 'metadata')]),
        metadata: 'not-an-object' as never,
      }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(
      controller.transact({
        ...transaction(controller, [command('command-1', 0, 'limited')]),
        commands: Array.from({ length: 1_001 }, (_, index) =>
          command(`command-${index}`, index, 'limited'),
        ),
      }),
    ).toMatchObject({ status: 'rejected', code: 'TRANSACTION_LIMIT_EXCEEDED' });
    expect(
      controller.transact(
        transaction(controller, [command('oversized', 0, 'x'.repeat(4 * 1_024 * 1_024))]),
      ),
    ).toMatchObject({ status: 'rejected', code: 'TRANSACTION_LIMIT_EXCEEDED' });
    expect(
      controller.transact(transaction(controller, [command('command-1', 0, 'gate')]), {
        permissionGate: () => {
          throw new Error('gate unavailable');
        },
      }),
    ).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_NOT_ALLOWED',
      message: 'gate unavailable',
    });

    const accessor = Object.defineProperty(
      {
        schemaVersion: 1 as const,
        id: 'accessor',
        baseRevision: 0,
        metadata: {},
      },
      'commands',
      {
        enumerable: true,
        get() {
          throw new Error('untrusted getter');
        },
      },
    ) as SerializableTransactionEnvelope;
    expect(() => controller.transact(accessor)).not.toThrow();
    expect(controller.transact(accessor)).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
    });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('isolates execute envelopes before reading command IDs or payloads', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const accessor = Object.defineProperty(
      { schemaVersion: 1 as const, command: command('nested', 0, 'value').command },
      'id',
      {
        enumerable: true,
        get() {
          throw new Error('command getter');
        },
      },
    ) as SerializableCommandEnvelope;

    expect(() => controller.execute(accessor)).not.toThrow();
    expect(controller.execute(accessor)).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
    });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('turns adversarial runtime inputs into stable rejection results', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    for (const input of [null, 1, 'transaction', [], { schemaVersion: 1, commands: [null] }]) {
      expect(() => controller.transact(input as never)).not.toThrow();
      expect(controller.transact(input as never)).toMatchObject({
        status: 'rejected',
        code: 'COMMAND_SCHEMA_INVALID',
      });
    }

    let reads = 0;
    const changing = Object.defineProperty(
      {
        schemaVersion: 1 as const,
        id: 'changing',
        baseRevision: 0,
        metadata: {},
      },
      'commands',
      {
        enumerable: true,
        get() {
          reads += 1;
          return reads === 1 ? [command('first', 0, 'first')] : [command('second', 0, 'second')];
        },
      },
    ) as SerializableTransactionEnvelope;
    expect(controller.transact(changing)).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
    });
    expect(reads).toBe(0);
    expect(
      controller.transact(transaction(controller, [command('command-1', 0, 'invalid-source')]), {
        source: 'remote' as never,
      }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('dry-runs against a candidate snapshot without changing document, history, or sequence', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = serializeSpreadsheetDocument(controller.getDocument());
    const listener = vi.fn();
    controller.subscribe(listener);

    const preview = controller.dryRun(
      transaction(controller, [command('command-1', 0, 'preview')]),
    );

    expect(preview.status).toBe('ready');
    if (preview.status === 'ready') {
      expect(preview.document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
        type: 'string',
        value: 'preview',
      });
    }
    expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(before);
    expect(controller.getSnapshot().revision).toBe(0);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
    expect(listener).not.toHaveBeenCalled();

    const committed = controller.execute(command('command-2', 0, 'accepted'));
    expect(committed.status).toBe('committed');
    if (committed.status === 'committed') {
      expect(committed.change.id).toMatch(/^change-\d+-1$/);
    }
  });

  it('rolls back every observable state field when a later command fails', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = controller.getSnapshot();
    const listener = vi.fn();
    controller.subscribe(listener);
    const invalid: SerializableCommandEnvelope = {
      schemaVersion: 1,
      id: 'invalid',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId('missing-sheet'), row: 0, column: 0 },
        text: 'invalid',
      },
    };

    expect(
      controller.transact(transaction(controller, [command('command-1', 0, 'temporary'), invalid])),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
    expect(listener).not.toHaveBeenCalled();

    const committed = controller.execute(command('command-2', 0, 'accepted'));
    expect(committed.status).toBe('committed');
    if (committed.status === 'committed') {
      expect(committed.change.id).toMatch(/^change-\d+-1$/);
    }
  });

  it('undoes and redoes a transaction as one semantically equivalent unit', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = serializeSpreadsheetDocument(controller.getDocument());
    controller.transact(
      transaction(controller, [command('command-1', 0, 'alpha'), command('command-2', 1, 'beta')]),
    );
    const committed = serializeSpreadsheetDocument(controller.getDocument());

    expect(controller.undo().status).toBe('committed');
    expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(before);
    expect(controller.redo().status).toBe('committed');
    expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(committed);
  });

  it('reports observer failure without throwing or rolling back an already committed transaction', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const observerError = new Error('observer failed');
    controller.subscribe(() => {
      throw observerError;
    });

    expect(
      controller.transact(transaction(controller, [command('command-1', 0, 'committed')])),
    ).toMatchObject({
      status: 'committed',
      notificationError: 'observer failed',
    });
    expect(controller.getSnapshot().revision).toBe(1);
    expect(controller.historySize).toEqual({ undo: 1, redo: 0 });
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 0, column: 0 })).toBe(
      'committed',
    );
  });
});
