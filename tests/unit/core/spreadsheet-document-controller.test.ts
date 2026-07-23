import { describe, expect, it, vi } from 'vitest';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import {
  cloneFrozenDocumentValue,
  SpreadsheetDocumentController,
} from '../../../src/core/controller/spreadsheet-document-controller';
import { WorkbookController } from '../../../src/core/controller/workbook-controller';
import { sheetId } from '../../../src/core';

function directDocument(overrides: Partial<SpreadsheetDocumentInput> = {}): SpreadsheetDocument {
  const input: SpreadsheetDocumentInput = {
    schemaVersion: 2,
    id: 'direct-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Direct',
          rowCount: 20,
          columnCount: 10,
          rows: [{ index: 4, height: 41, hidden: true }],
          columns: [{ index: 3, width: 123, hidden: true }],
          cells: [
            {
              row: 2,
              column: 3,
              cell: {
                input: {
                  type: 'custom',
                  cellType: 'acme.widget',
                  schemaVersion: 7,
                  value: { payload: 'kept' },
                },
                resourceId: 'resource-1',
                templateId: 'template-1',
                styleId: 'style-1',
                metadata: { untouched: true },
                editable: true,
                printable: false,
              },
            },
          ],
          merges: [{ start: { row: 5, column: 1 }, end: { row: 6, column: 2 } }],
          freeze: { row: 2, column: 1 },
          filter: {
            range: { start: { row: 0, column: 0 }, end: { row: 10, column: 4 } },
            filters: [{ column: 0, operator: 'in', values: ['kept'] }],
            sort: { column: 1, direction: 'desc' },
          },
        },
      ],
      styles: [{ id: 'style-1', value: { color: '#123456' } }],
      validations: [],
      settings: { dateSystem: 'excel-1904', localeHint: 'en-GB' },
    },
    templates: [
      {
        id: 'template-1',
        name: 'Template',
        sheetId: 'sheet-1',
        printProfile: {
          paperSize: 'A4',
          orientation: 'portrait',
          margins: { top: 1, right: 1, bottom: 1, left: 1 },
        },
      },
    ],
    resources: {
      items: [
        {
          id: 'resource-1',
          kind: 'image',
          metadata: { source: 'direct' },
        },
      ],
    },
    extensions: { 'acme.runtime': { enabled: true } },
    ...overrides,
  };
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) throw new Error('Direct schema 2 fixture must be valid');
  return parsed.document;
}

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

  it('preserves every untouched schema 2 field through constructor, dispatch, and replace', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    expect(controller.getDocument()).toEqual(initial);

    controller.dispatch(
      { type: 'rename-sheet', sheet: sheetId('sheet-1'), name: 'Renamed' },
      'ref',
    );
    expect(controller.getDocument()).toEqual({
      ...initial,
      workbook: {
        ...initial.workbook,
        sheets: [{ ...initial.workbook.sheets[0]!, name: 'Renamed' }],
      },
    });

    const replacement = directDocument({ id: 'replacement-document' });
    controller.replace(replacement);
    expect(controller.getDocument()).toEqual(replacement);
  });

  it('keeps shared validation references sheet-local without creating cells', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'validation-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'One',
            cells: [
              {
                row: 0,
                column: 0,
                cell: { input: { type: 'string', value: 'one' }, validationId: 'validation-1' },
              },
            ],
            merges: [],
          },
          {
            id: 'sheet-2',
            name: 'Two',
            cells: [
              {
                row: 0,
                column: 1,
                cell: { input: { type: 'string', value: 'two' }, validationId: 'validation-1' },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [
          {
            id: 'validation-1',
            value: {
              mode: 'cell',
              type: 'number',
              required: false,
              operator: 'gte',
              value: '0',
            },
          },
        ],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error('Validation fixture must be valid');

    const controller = new SpreadsheetDocumentController(parsed.document);
    expect(controller.getDocument().workbook.sheets.map((sheet) => sheet.cells)).toEqual(
      parsed.document.workbook.sheets.map((sheet) => sheet.cells),
    );
  });

  it('clears operational style fields without degrading untouched custom cell input', () => {
    const controller = new SpreadsheetDocumentController(directDocument());
    controller.dispatch(
      {
        type: 'clear-format',
        selection: {
          sheet: sheetId('sheet-1'),
          active: { row: 2, column: 3 },
          range: { start: { row: 2, column: 3 }, end: { row: 2, column: 3 } },
        },
      },
      'toolbar',
    );

    const cell = controller.getDocument().workbook.sheets[0]?.cells[0]?.cell;
    expect(cell?.styleId).toBeUndefined();
    expect(cell?.input).toEqual({
      type: 'custom',
      cellType: 'acme.widget',
      schemaVersion: 7,
      value: { payload: 'kept' },
    });
    expect(cell?.metadata).toEqual({ untouched: true });
  });

  it('rolls back document, projection, revision, history, sequencing, and notifications atomically', () => {
    const controller = new SpreadsheetDocumentController(directDocument());
    const before = controller.getSnapshot();
    const beforeHistory = controller.historySize;
    const subscriber = vi.fn();
    controller.subscribe(subscriber);

    expect(() =>
      controller.dispatch({ type: 'delete-sheet', sheet: sheetId('sheet-1') }, 'ref', {
        beforeNotify: () => undefined,
      }),
    ).toThrow();

    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual(beforeHistory);
    expect(subscriber).not.toHaveBeenCalled();
  });

  it('does not expose a replacement document when the operation projection rejects it', () => {
    const controller = new SpreadsheetDocumentController(directDocument());
    const before = controller.getSnapshot();
    const beforeHistory = controller.historySize;
    vi.spyOn(WorkbookController.prototype, 'replace').mockImplementationOnce(() => {
      throw new Error('projection rejected');
    });

    expect(() => controller.replace(directDocument({ id: 'rejected-replacement' }))).toThrow(
      'projection rejected',
    );
    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual(beforeHistory);
  });

  it('restores change sequencing when beforeNotify rejects a commit', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    expect(() =>
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
          text: 'rejected',
        },
        'ref',
        {
          beforeNotify: () => {
            throw new Error('reject checkpoint');
          },
        },
      ),
    ).toThrow('reject checkpoint');

    const accepted = controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'accepted',
      },
      'ref',
    );
    expect(accepted.status).toBe('committed');
    if (accepted.status === 'committed')
      expect(accepted.commit.change.id).toMatch(/^change-\d+-1$/);
  });

  it('clones dangerous JSON keys into frozen null-prototype records', () => {
    const dangerous = JSON.parse(
      '{"__proto__":{"polluted":true},"constructor":{"prototype":{"polluted":true}}}',
    ) as Record<string, unknown>;
    const snapshot = cloneFrozenDocumentValue(dangerous);
    expect(Object.getPrototypeOf(snapshot)).toBeNull();
    expect(Object.hasOwn(snapshot, '__proto__')).toBe(true);
    expect(Object.hasOwn(snapshot, 'constructor')).toBe(true);
    expect(({} as { polluted?: boolean }).polluted).toBeUndefined();
  });
});
