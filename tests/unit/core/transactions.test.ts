import { describe, expect, it, vi } from 'vitest';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
} from '../../../src/document';
import { sheetId } from '../../../src/core';
import { getCellData } from '../../../src/core/model/cells';
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

function createMultiSheetDocument() {
  const input = JSON.parse(
    serializeSpreadsheetDocument(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    ),
  ) as {
    workbook: {
      sheets: Array<{
        id: string;
        name: string;
        cells: unknown[];
        merges: unknown[];
      }>;
    };
  };
  input.workbook.sheets.push({
    id: 'sheet-2',
    name: 'Second',
    cells: [],
    merges: [],
  });
  const parsed = parseSpreadsheetDocument(input as never);
  if (!parsed.ok) throw new TypeError('Expected test document to parse');
  return parsed.document;
}

function createFormulaDocument() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'formula-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: [
            { row: 0, column: 0, cell: { input: { type: 'number', value: 1 } } },
            { row: 0, column: 1, cell: { input: { type: 'formula', source: '=A1+1' } } },
          ],
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
  if (!parsed.ok) throw new TypeError('Expected formula document to parse');
  return parsed.document;
}

function createCrossSheetFormulaDocument() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'cross-sheet-formula-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Inputs',
          cells: [{ row: 0, column: 0, cell: { input: { type: 'number', value: 2 } } }],
          merges: [],
        },
        {
          id: 'sheet-2',
          name: 'Results',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'formula', source: '=Inputs!A1*2' } },
            },
          ],
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
  if (!parsed.ok) throw new TypeError('Expected cross-sheet formula document to parse');
  return parsed.document;
}

function createBlankFormulaDocument() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'blank-formula-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'formula', source: '=IF(FALSE,1)' } },
            },
          ],
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
  if (!parsed.ok) throw new TypeError('Expected blank formula document to parse');
  return parsed.document;
}

function createLegacySyntaxFormulaDocument() {
  const input = structuredClone(createCrossSheetFormulaDocument());
  const formula = input.workbook.sheets[1]?.cells[0]?.cell.input;
  if (formula?.type !== 'formula') throw new TypeError('Expected formula fixture');
  (formula as { source: string }).source = '=SUM(Inputs!A:A)';
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) throw new TypeError('Expected legacy-syntax formula document to parse');
  return parsed.document;
}

function projectedValue(
  controller: SpreadsheetDocumentController,
  row: number,
  column: number,
  sheetIndex = 0,
) {
  const sheet = controller.getSnapshot().projection[sheetIndex];
  return sheet === undefined ? undefined : getCellData(sheet, row, column)?.value;
}

describe('SpreadsheetDocumentController transactions', () => {
  it('owns revision-scoped typed calculation through edit, undo, and redo', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());
    expect(projectedValue(controller, 0, 1)).toBe(2);

    expect(controller.execute(command('edit-input', 0, '5')).status).toBe('committed');
    expect(projectedValue(controller, 0, 1)).toBe(6);

    expect(controller.undo().status).toBe('committed');
    expect(projectedValue(controller, 0, 1)).toBe(2);
    expect(controller.redo().status).toBe('committed');
    expect(projectedValue(controller, 0, 1)).toBe(6);
  });

  it('recalculates cross-sheet dependents after a source-cell edit', () => {
    const controller = new SpreadsheetDocumentController(createCrossSheetFormulaDocument());
    expect(projectedValue(controller, 0, 0, 1)).toBe(4);

    expect(controller.execute(command('edit-cross-sheet-input', 0, '7')).status).toBe('committed');

    expect(projectedValue(controller, 0, 0, 1)).toBe(14);
  });

  it('recompiles formulas after structural edits and history traversal', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());

    expect(
      controller.dispatch(
        { type: 'insert-row', sheet: sheetId('sheet-1'), index: 0 },
        'context-menu',
      ).status,
    ).toBe('committed');
    expect(
      controller.getDocument().workbook.sheets[0]?.cells.find((cell) => cell.column === 1)?.cell
        .input,
    ).toEqual({ type: 'formula', source: '=A2+1' });
    expect(projectedValue(controller, 1, 1)).toBe(2);

    expect(controller.undo().status).toBe('committed');
    expect(projectedValue(controller, 0, 1)).toBe(2);
  });

  it('preserves cross-sheet formulas when the referenced sheet is renamed', () => {
    const controller = new SpreadsheetDocumentController(createCrossSheetFormulaDocument());

    expect(
      controller.dispatch(
        { type: 'rename-sheet', sheet: sheetId('sheet-1'), name: 'Renamed Inputs' },
        'sheet-tabs',
      ).status,
    ).toBe('committed');

    expect(controller.getDocument().workbook.sheets[1]?.cells[0]?.cell.input).toEqual({
      type: 'formula',
      source: "='Renamed Inputs'!A1*2",
    });
    expect(projectedValue(controller, 0, 0, 1)).toBe(4);
  });

  it('renames sheet qualifiers in formulas that require the legacy syntax fallback', () => {
    const controller = new SpreadsheetDocumentController(createLegacySyntaxFormulaDocument());

    expect(
      controller.dispatch(
        { type: 'rename-sheet', sheet: sheetId('sheet-1'), name: 'Renamed Inputs' },
        'sheet-tabs',
      ).status,
    ).toBe('committed');

    expect(controller.getDocument().workbook.sheets[1]?.cells[0]?.cell.input).toEqual({
      type: 'formula',
      source: "=SUM('Renamed Inputs'!A:A)",
    });
  });

  it('rolls back a direct dispatch when formula dependency compilation exceeds its budget', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());
    const before = controller.getSnapshot();

    expect(() =>
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 1 },
          text: '=SUM(A1:A100001)',
        },
        'ref',
      ),
    ).toThrow(/dependency limit/u);

    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
    expect(projectedValue(controller, 0, 1)).toBe(2);
  });

  it('rebuilds typed calculation state when the document is replaced', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());

    controller.replace(createCrossSheetFormulaDocument());

    expect(projectedValue(controller, 0, 0, 1)).toBe(4);
    expect(controller.execute(command('edit-replaced-input', 0, '3')).status).toBe('committed');
    expect(projectedValue(controller, 0, 0, 1)).toBe(6);
  });

  it('projects typed formula blanks as an explicit cache value', () => {
    const controller = new SpreadsheetDocumentController(createBlankFormulaDocument());

    expect(projectedValue(controller, 0, 0)).toBeNull();
  });

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

  it('returns an immutable committed record with forward, inverse, and diagnostic data', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );

    const outcome = controller.transact(
      transaction(controller, [command('command-1', 0, 'alpha')]),
    );

    expect(outcome.status).toBe('committed');
    if (outcome.status !== 'committed') return;
    expect(outcome.transaction).toMatchObject({
      id: 'transaction-1',
      committedRevision: 1,
      diagnostics: [],
    });
    expect(outcome.transaction.forwardPatches.length).toBeGreaterThan(0);
    expect(outcome.transaction.inversePatches.length).toBeGreaterThan(0);
    expect(JSON.parse(JSON.stringify(outcome.transaction))).toEqual(outcome.transaction);
    expect(Object.isFrozen(outcome.transaction)).toBe(true);
    expect(Object.isFrozen(outcome.transaction.forwardPatches)).toBe(true);
    expect(Object.isFrozen(outcome.transaction.inversePatches)).toBe(true);
  });

  it('aggregates every affected sheet and range in one transaction change', () => {
    const controller = new SpreadsheetDocumentController(createMultiSheetDocument());
    const listener = vi.fn();
    controller.subscribe(listener);
    const second = {
      ...command('command-2', 3, 'beta'),
      command: {
        ...command('command-2', 3, 'beta').command,
        address: { sheet: sheetId('sheet-2'), row: 3, column: 0 },
      },
    } as SerializableCommandEnvelope;

    const outcome = controller.transact(
      transaction(controller, [command('command-1', 1, 'alpha'), second]),
    );

    expect(outcome.status).toBe('committed');
    if (outcome.status !== 'committed') return;
    expect(outcome.change.aggregate).toEqual({
      commandCount: 2,
      sheets: [
        {
          sheet: sheetId('sheet-1'),
          kinds: ['cell'],
          ranges: [
            {
              start: { row: 1, column: 0 },
              end: { row: 1, column: 0 },
            },
          ],
        },
        {
          sheet: sheetId('sheet-2'),
          kinds: ['cell'],
          ranges: [
            {
              start: { row: 3, column: 0 },
              end: { row: 3, column: 0 },
            },
          ],
        },
      ],
    });
    listener.mockClear();
    expect(controller.undo().status).toBe('committed');
    expect(listener.mock.calls[0]?.[0].commit.change).toMatchObject({
      kind: 'history',
      aggregate: outcome.change.aggregate,
    });
    listener.mockClear();
    expect(controller.redo().status).toBe('committed');
    expect(listener.mock.calls[0]?.[0].commit.change).toMatchObject({
      kind: 'history',
      aggregate: outcome.change.aggregate,
    });
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

    expect(outcome).toMatchObject({
      status: 'rejected',
      code: 'COMMAND_NOT_ALLOWED',
    });
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
      {
        schemaVersion: 1 as const,
        command: command('nested', 0, 'value').command,
      },
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

  it('restores formula calculation state after a dry-run', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());

    expect(
      controller.dryRun(transaction(controller, [command('preview-input', 0, '5')])).status,
    ).toBe('ready');
    expect(
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 2 },
          text: 'unrelated',
        },
        'ref',
      ).status,
    ).toBe('committed');

    expect(projectedValue(controller, 0, 1)).toBe(2);
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

  it('restores formula calculation state after a failed multi-command transaction', () => {
    const controller = new SpreadsheetDocumentController(createFormulaDocument());
    const invalid: SerializableCommandEnvelope = {
      schemaVersion: 1,
      id: 'invalid-formula-transaction-command',
      command: {
        type: 'set-cell-text',
        address: { sheet: sheetId('missing-sheet'), row: 0, column: 0 },
        text: 'invalid',
      },
    };

    expect(
      controller.transact(
        transaction(controller, [command('temporary-formula-input', 0, '5'), invalid]),
      ),
    ).toMatchObject({ status: 'rejected' });
    expect(
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 2 },
          text: 'unrelated',
        },
        'ref',
      ).status,
    ).toBe('committed');

    expect(projectedValue(controller, 0, 1)).toBe(2);
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

  it('reports observer failures after direct dispatch, undo, and redo without throwing', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.execute(command('initial', 0, 'initial'));
    controller.subscribe(() => {
      throw new Error('direct observer failed');
    });

    expect(controller.dispatch(command('direct', 1, 'direct').command, 'ref')).toMatchObject({
      status: 'committed',
      commit: { notificationError: 'direct observer failed' },
    });
    expect(controller.undo()).toMatchObject({
      status: 'committed',
      commit: { notificationError: 'direct observer failed' },
    });
    expect(controller.redo()).toMatchObject({
      status: 'committed',
      commit: { notificationError: 'direct observer failed' },
    });
  });
});
