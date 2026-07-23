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
import { History } from '../../../src/core/controller/history';
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
              row: 0,
              column: 0,
              cell: {
                input: {
                  type: 'custom',
                  cellType: 'acme.widget',
                  schemaVersion: 7,
                  value: { payload: 'before' },
                },
                resourceId: 'resource-1',
                templateId: 'template-1',
                metadata: { position: 'before' },
              },
            },
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
                validationId: 'validation-1',
                metadata: { untouched: true },
                editable: true,
                printable: false,
              },
            },
            {
              row: 5,
              column: 6,
              cell: {
                input: {
                  type: 'custom',
                  cellType: 'acme.widget',
                  schemaVersion: 7,
                  value: { payload: 'after' },
                },
                resourceId: 'resource-1',
                templateId: 'template-1',
                metadata: { position: 'after' },
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
      settings: { dateSystem: 'excel-1904', localeHint: 'en-GB' },
    },
    templates: [
      {
        id: 'template-1',
        name: 'Template',
        bindings: [
          {
            id: 'conditional-1' as never,
            type: 'conditional-range',
            range: {
              sheetId: 'sheet-1' as never,
              start: { row: 2, column: 3 },
              end: { row: 5, column: 6 },
            },
            when: 'visible',
          },
        ],
        printProfiles: [],
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

function documentWithEveryInput(): SpreadsheetDocument {
  const inputs: SpreadsheetDocument['workbook']['sheets'][number]['cells'][number]['cell']['input'][] =
    [
      { type: 'string', value: 'plain' },
      { type: 'number', value: 7 },
      { type: 'boolean', value: true },
      { type: 'formula', source: '=A1' },
      {
        type: 'custom',
        cellType: 'acme.all-inputs',
        schemaVersion: 1,
        value: { nested: true },
      },
    ];
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'all-inputs',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Inputs',
          cells: inputs.map((input, column) => ({
            row: 0,
            column,
            cell: { input, metadata: { inputType: input.type } },
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
  if (!parsed.ok) throw new Error('All-input fixture must be valid');
  return parsed.document;
}

function schemaOnlyClipboardDocument(sameTarget = false): SpreadsheetDocument {
  const source = {
    input: {
      type: 'custom' as const,
      cellType: 'acme.source',
      schemaVersion: 1,
      value: { legacy: 'same' },
    },
    metadata: { provenance: 'source' },
  };
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'schema-only-clipboard',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Schema only',
          cells: [
            { row: 0, column: 0, cell: source },
            {
              row: 0,
              column: 1,
              cell: sameTarget
                ? source
                : {
                    input: {
                      type: 'custom',
                      cellType: 'acme.target',
                      schemaVersion: 9,
                      value: { legacy: 'same' },
                    },
                    metadata: { provenance: 'target' },
                  },
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
  if (!parsed.ok) throw new Error('Schema-only clipboard fixture must be valid');
  return parsed.document;
}

function sparseCutDocument(): SpreadsheetDocument {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'sparse-cut',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sparse cut',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'string', value: 'move' } },
            },
            {
              row: 1,
              column: 1,
              cell: {
                input: {
                  type: 'custom',
                  cellType: 'acme.rich-target',
                  schemaVersion: 3,
                  value: { keep: true },
                },
                metadata: { untouched: true },
              },
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
  if (!parsed.ok) throw new Error('Sparse cut fixture must be valid');
  return parsed.document;
}

function selection(
  sheet: ReturnType<typeof sheetId>,
  row: number,
  column: number,
  endRow = row,
  endColumn = column,
) {
  return {
    sheet,
    active: { row, column },
    range: {
      start: { row, column },
      end: { row: endRow, column: endColumn },
    },
  } as const;
}

function richCells(document: SpreadsheetDocument) {
  return document.workbook.sheets[0]!.cells.filter(
    ({ cell }) =>
      cell.resourceId !== undefined || cell.templateId !== undefined || cell.metadata !== undefined,
  ).map(({ row, column, cell }) => ({
    row,
    column,
    cell,
  }));
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

    const cell = controller
      .getDocument()
      .workbook.sheets[0]?.cells.find((item) => item.row === 2 && item.column === 3)?.cell;
    expect(cell?.styleId).toBeUndefined();
    expect(cell?.input).toEqual({
      type: 'custom',
      cellType: 'acme.widget',
      schemaVersion: 7,
      value: { payload: 'kept' },
    });
    expect(cell?.metadata).toEqual({ untouched: true });
  });

  it.each([
    {
      command: { type: 'insert-row', sheet: sheetId('sheet-1'), index: 2, count: 2 } as const,
      coordinates: [
        [0, 0],
        [4, 3],
        [7, 6],
      ],
      layoutIndex: 6,
      templateRange: {
        sheetId: 'sheet-1',
        start: { row: 4, column: 3 },
        end: { row: 7, column: 6 },
      },
    },
    {
      command: { type: 'delete-row', sheet: sheetId('sheet-1'), index: 1, count: 2 } as const,
      coordinates: [
        [0, 0],
        [3, 6],
      ],
      layoutIndex: 2,
      templateRange: {
        sheetId: 'sheet-1',
        start: { row: 1, column: 3 },
        end: { row: 3, column: 6 },
      },
    },
    {
      command: { type: 'insert-column', sheet: sheetId('sheet-1'), index: 3, count: 2 } as const,
      coordinates: [
        [0, 0],
        [2, 5],
        [5, 8],
      ],
      layoutIndex: 5,
      templateRange: {
        sheetId: 'sheet-1',
        start: { row: 2, column: 5 },
        end: { row: 5, column: 8 },
      },
    },
    {
      command: { type: 'delete-column', sheet: sheetId('sheet-1'), index: 3, count: 1 } as const,
      coordinates: [
        [0, 0],
        [5, 5],
      ],
      layoutIndex: null,
      templateRange: {
        sheetId: 'sheet-1',
        start: { row: 2, column: 3 },
        end: { row: 5, column: 5 },
      },
    },
  ])(
    'transforms complete rich cells for $command.type',
    ({ command, coordinates, layoutIndex, templateRange }) => {
      const controller = new SpreadsheetDocumentController(directDocument());
      controller.dispatch(command, 'context-menu');

      const document = controller.getDocument();
      const cells = richCells(document);
      expect(cells.map((cell) => [cell.row, cell.column])).toEqual(coordinates);
      for (const item of cells) {
        expect(item.cell.input.type).toBe('custom');
        expect(item.cell.resourceId).toBe('resource-1');
        expect(item.cell.templateId).toBe('template-1');
        expect(item.cell.metadata).toBeDefined();
      }
      const middle = cells.find(
        (item) => JSON.stringify(item.cell.metadata) === '{"untouched":true}',
      );
      if (middle !== undefined) expect(middle.cell.validationId).toBe('validation-1');
      const layout = command.type.endsWith('row')
        ? document.workbook.sheets[0]?.rows
        : document.workbook.sheets[0]?.columns;
      expect(layout?.map((item) => item.index)).toEqual(layoutIndex === null ? [] : [layoutIndex]);
      expect(document.templates[0]?.bindings[0]).toMatchObject({
        type: 'conditional-range',
        range: templateRange,
      });
    },
  );

  it('restores and reapplies rich structural transforms through history', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const command = { type: 'insert-row', sheet: sheetId('sheet-1'), index: 2, count: 2 } as const;
    controller.dispatch(command, 'context-menu');
    const inserted = controller.getDocument();

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    controller.redo();
    expect(controller.getDocument()).toEqual(inserted);
  });

  it('keeps duplicate-valued rich cells distinct when sorting', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    controller.dispatch(
      { type: 'sort', sheet: sheetId('sheet-1'), column: 0, order: 'asc' },
      'toolbar',
    );

    expect(richCells(controller.getDocument())).toEqual(richCells(initial));
  });

  it.each([
    { cut: false, label: 'copy' },
    { cut: true, label: 'cut' },
  ])('$label preserves the complete rich source cell', ({ cut }) => {
    const controller = new SpreadsheetDocumentController(directDocument());
    const sheet = sheetId('sheet-1');
    const source = richCells(controller.getDocument()).find(
      (item) => item.row === 2 && item.column === 3,
    )!.cell;
    controller.dispatch(
      {
        type: 'paste-internal',
        source: selection(sheet, 2, 3),
        target: selection(sheet, 8, 8),
        mode: 'all',
        cut,
      },
      'context-menu',
    );

    const cells = richCells(controller.getDocument());
    expect(cells.find((item) => item.row === 8 && item.column === 8)?.cell).toEqual(source);
    expect(cells.some((item) => item.row === 2 && item.column === 3)).toBe(!cut);
  });

  it('removes target validation when all-mode paste copies an unvalidated cell', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');
    const validationRefs = () =>
      controller.getSnapshot().projection[0]?.validations?.flatMap((rule) => rule.refs ?? []) ?? [];
    const command = {
      type: 'paste-internal',
      source: selection(sheet, 0, 0),
      target: selection(sheet, 2, 3),
      mode: 'all',
      cut: false,
    } as const;

    controller.dispatch(command, 'context-menu');
    const pasted = controller.getDocument();
    expect(
      pasted.workbook.sheets[0]?.cells.find((item) => item.row === 2 && item.column === 3)?.cell
        .validationId,
    ).toBeUndefined();
    expect(validationRefs()).not.toContain('D3');

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    expect(validationRefs()).toContain('D3');
    controller.redo();
    expect(controller.getDocument()).toEqual(pasted);
    expect(validationRefs()).not.toContain('D3');

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet, row: 2, column: 3 },
        text: '-1',
      },
      'ref',
    );
    expect(
      controller
        .getDocument()
        .workbook.sheets[0]?.cells.find((item) => item.row === 2 && item.column === 3)?.cell
        .validationId,
    ).toBeUndefined();
    expect(validationRefs()).not.toContain('D3');
    expect(controller.validate()).toEqual({ valid: true, issues: [] });
  });

  it('keeps copied validation aligned across document, projection, validation, edits, and history', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');
    const validationRefs = () =>
      controller.getSnapshot().projection[0]?.validations?.flatMap((rule) => rule.refs ?? []) ?? [];
    const command = {
      type: 'paste-internal',
      source: selection(sheet, 2, 3),
      target: selection(sheet, 8, 8),
      mode: 'all',
      cut: false,
    } as const;

    controller.dispatch(command, 'context-menu');
    const pasted = controller.getDocument();
    expect(
      pasted.workbook.sheets[0]?.cells.find((item) => item.row === 8 && item.column === 8)?.cell
        .validationId,
    ).toBe('validation-1');
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'I9']));

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    expect(validationRefs()).toEqual(['D3']);
    controller.redo();
    expect(controller.getDocument()).toEqual(pasted);
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'I9']));

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet, row: 8, column: 8 },
        text: '-1',
      },
      'ref',
    );
    expect(
      controller
        .getDocument()
        .workbook.sheets[0]?.cells.find((item) => item.row === 8 && item.column === 8)?.cell
        .validationId,
    ).toBe('validation-1');
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'I9']));
    expect(controller.validate()).toMatchObject({
      valid: false,
      issues: [{ address: { sheet, row: 8, column: 8 } }],
    });
  });

  it('removes a validated cut source without leaving a blank validation-only cell', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');
    const command = {
      type: 'paste-internal',
      source: selection(sheet, 2, 3),
      target: selection(sheet, 8, 8),
      mode: 'all',
      cut: true,
    } as const;

    controller.dispatch(command, 'context-menu');
    const cut = controller.getDocument();
    const cells = cut.workbook.sheets[0]!.cells;
    expect(cells.find((item) => item.row === 2 && item.column === 3)).toBeUndefined();
    expect(cells.find((item) => item.row === 8 && item.column === 8)?.cell.validationId).toBe(
      'validation-1',
    );

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    controller.redo();
    expect(controller.getDocument()).toEqual(cut);
  });

  it('autofill preserves the complete rich source cell', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');
    const validationRefs = () =>
      controller.getSnapshot().projection[0]?.validations?.flatMap((rule) => rule.refs ?? []) ?? [];
    const source = richCells(controller.getDocument()).find(
      (item) => item.row === 2 && item.column === 3,
    )!.cell;
    controller.dispatch(
      {
        type: 'autofill',
        source: selection(sheet, 2, 3),
        target: selection(sheet, 8, 3),
        mode: 'all',
      },
      'pointer',
    );

    const filled = controller.getDocument();
    expect(richCells(filled).find((item) => item.row === 8 && item.column === 3)?.cell).toEqual(
      source,
    );
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'D9']));

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    expect(validationRefs()).toEqual(['D3']);
    controller.redo();
    expect(controller.getDocument()).toEqual(filled);
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'D9']));

    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet, row: 8, column: 3 },
        text: '-1',
      },
      'ref',
    );
    expect(
      controller
        .getDocument()
        .workbook.sheets[0]?.cells.find((item) => item.row === 8 && item.column === 3)?.cell
        .validationId,
    ).toBe('validation-1');
    expect(validationRefs()).toEqual(expect.arrayContaining(['D3', 'D9']));
    expect(controller.validate()).toMatchObject({
      valid: false,
      issues: [{ address: { sheet, row: 8, column: 3 } }],
    });
  });

  it('removes target validation when all-mode autofill uses an unvalidated cell', () => {
    const initial = directDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');
    const validationRefs = () =>
      controller.getSnapshot().projection[0]?.validations?.flatMap((rule) => rule.refs ?? []) ?? [];
    const command = {
      type: 'autofill',
      source: selection(sheet, 0, 0),
      target: selection(sheet, 2, 3),
      mode: 'all',
    } as const;

    controller.dispatch(command, 'pointer');
    const filled = controller.getDocument();
    expect(
      filled.workbook.sheets[0]?.cells.find((item) => item.row === 2 && item.column === 3)?.cell
        .validationId,
    ).toBeUndefined();
    expect(validationRefs()).not.toContain('D3');

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    expect(validationRefs()).toContain('D3');
    controller.redo();
    expect(controller.getDocument()).toEqual(filled);
    expect(validationRefs()).not.toContain('D3');
  });

  it.each([
    {
      label: 'copy',
      command: (sheet: ReturnType<typeof sheetId>) =>
        ({
          type: 'paste-internal',
          source: selection(sheet, 0, 0),
          target: selection(sheet, 0, 1),
          mode: 'all',
          cut: false,
        }) as const,
    },
    {
      label: 'autofill',
      command: (sheet: ReturnType<typeof sheetId>) =>
        ({
          type: 'autofill',
          source: selection(sheet, 0, 0),
          target: selection(sheet, 0, 1),
          mode: 'all',
        }) as const,
    },
  ])('commits a schema-only $label exactly once with history', ({ command }) => {
    const initial = schemaOnlyClipboardDocument();
    const source = initial.workbook.sheets[0]!.cells[0]!.cell;
    const controller = new SpreadsheetDocumentController(initial);
    const subscriber = vi.fn();
    const beforeNotify = vi.fn();
    controller.subscribe(subscriber);

    const outcome = controller.dispatch(command(sheetId('sheet-1')), 'context-menu', {
      beforeNotify,
    });

    expect(outcome.status).toBe('committed');
    expect(controller.getDocument().workbook.sheets[0]?.cells[1]?.cell).toEqual(source);
    expect(controller.getSnapshot().revision).toBe(1);
    expect(controller.historySize).toEqual({ undo: 1, redo: 0 });
    expect(beforeNotify).toHaveBeenCalledOnce();
    expect(subscriber).toHaveBeenCalledOnce();

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    controller.redo();
    expect(controller.getDocument().workbook.sheets[0]?.cells[1]?.cell).toEqual(source);
  });

  it.each([
    {
      label: 'copy',
      command: (sheet: ReturnType<typeof sheetId>) =>
        ({
          type: 'paste-internal',
          source: selection(sheet, 0, 0),
          target: selection(sheet, 0, 1),
          mode: 'all',
          cut: false,
        }) as const,
    },
    {
      label: 'autofill',
      command: (sheet: ReturnType<typeof sheetId>) =>
        ({
          type: 'autofill',
          source: selection(sheet, 0, 0),
          target: selection(sheet, 0, 1),
          mode: 'all',
        }) as const,
    },
  ])('keeps a complete schema and operational $label no-op silent', ({ command }) => {
    const controller = new SpreadsheetDocumentController(schemaOnlyClipboardDocument(true));
    const before = controller.getSnapshot();
    const subscriber = vi.fn();
    const beforeNotify = vi.fn();
    controller.subscribe(subscriber);

    expect(
      controller.dispatch(command(sheetId('sheet-1')), 'context-menu', { beforeNotify }),
    ).toEqual({ status: 'noop' });
    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
    expect(beforeNotify).not.toHaveBeenCalled();
    expect(subscriber).not.toHaveBeenCalled();
  });

  it('rolls back a rejected schema-only projection commit atomically', () => {
    const initial = schemaOnlyClipboardDocument();
    const controller = new SpreadsheetDocumentController(initial);
    const before = controller.getSnapshot();
    const subscriber = vi.fn();
    controller.subscribe(subscriber);
    const command = {
      type: 'paste-internal',
      source: selection(sheetId('sheet-1'), 0, 0),
      target: selection(sheetId('sheet-1'), 0, 1),
      mode: 'all',
      cut: false,
    } as const;

    expect(() =>
      controller.dispatch(command, 'context-menu', {
        beforeNotify: () => {
          throw new Error('reject schema-only projection');
        },
      }),
    ).toThrow('reject schema-only projection');
    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.getDocument()).toEqual(initial);
    expect(controller.historySize).toEqual({ undo: 0, redo: 0 });
    expect(subscriber).not.toHaveBeenCalled();

    const accepted = controller.dispatch(command, 'context-menu');
    expect(accepted.status).toBe('committed');
    if (accepted.status === 'committed') {
      expect(accepted.commit.change.id).toMatch(/^change-\d+-1$/);
    }
  });

  it('keeps rich targets aligned with sparse cut semantics through undo and redo', () => {
    const initial = sparseCutDocument();
    const richTarget = initial.workbook.sheets[0]!.cells.find(
      (item) => item.row === 1 && item.column === 1,
    )!.cell;
    const controller = new SpreadsheetDocumentController(initial);
    const sheet = sheetId('sheet-1');

    controller.dispatch(
      {
        type: 'paste-internal',
        source: selection(sheet, 0, 0, 0, 1),
        target: selection(sheet, 1, 0, 1, 1),
        mode: 'all',
        cut: true,
      },
      'context-menu',
    );
    const cut = controller.getDocument();
    const cells = cut.workbook.sheets[0]!.cells;
    expect(cells.find((item) => item.row === 0 && item.column === 0)).toBeUndefined();
    expect(cells.find((item) => item.row === 1 && item.column === 0)?.cell.input).toEqual({
      type: 'string',
      value: 'move',
    });
    expect(cells.find((item) => item.row === 1 && item.column === 1)?.cell).toEqual(richTarget);

    controller.undo();
    expect(controller.getDocument()).toEqual(initial);
    controller.redo();
    expect(controller.getDocument()).toEqual(cut);
  });

  it('copies every CellInput variant with complete cell metadata', () => {
    const controller = new SpreadsheetDocumentController(documentWithEveryInput());
    const sheet = sheetId('sheet-1');
    const sources = [...controller.getDocument().workbook.sheets[0]!.cells];
    for (let column = 0; column < sources.length; column += 1) {
      controller.dispatch(
        {
          type: 'paste-internal',
          source: selection(sheet, 0, column),
          target: selection(sheet, 1, column),
          mode: 'all',
          cut: false,
        },
        'context-menu',
      );
    }
    const cells = controller.getDocument().workbook.sheets[0]!.cells;
    for (let column = 0; column < sources.length; column += 1) {
      expect(cells.find((item) => item.row === 1 && item.column === column)?.cell).toEqual(
        sources[column]?.cell,
      );
    }
  });

  it('autofills every CellInput variant while preserving complete cell metadata', () => {
    const controller = new SpreadsheetDocumentController(documentWithEveryInput());
    const sheet = sheetId('sheet-1');
    for (let column = 0; column < 5; column += 1) {
      controller.dispatch(
        {
          type: 'autofill',
          source: selection(sheet, 0, column),
          target: selection(sheet, 1, column),
          mode: 'all',
        },
        'pointer',
      );
    }
    const cells = controller.getDocument().workbook.sheets[0]!.cells;
    const expectedInputs = [
      { type: 'string', value: 'plain' },
      { type: 'number', value: 7 },
      { type: 'boolean', value: true },
      { type: 'formula', source: '=A2' },
      {
        type: 'custom',
        cellType: 'acme.all-inputs',
        schemaVersion: 1,
        value: { nested: true },
      },
    ];
    for (let column = 0; column < expectedInputs.length; column += 1) {
      expect(cells.find((item) => item.row === 1 && item.column === column)?.cell).toEqual({
        input: expectedInputs[column],
        metadata: { inputType: expectedInputs[column]?.type },
      });
    }
  });

  it('restores custom input and references through edit undo redo', () => {
    const controller = new SpreadsheetDocumentController(directDocument());
    const address = { sheet: sheetId('sheet-1'), row: 2, column: 3 };
    const original = richCells(controller.getDocument()).find(
      (item) => item.row === 2 && item.column === 3,
    )!.cell;

    controller.dispatch({ type: 'set-cell-text', address, text: 'edited' }, 'ref');
    controller.undo();
    expect(
      richCells(controller.getDocument()).find((item) => item.row === 2 && item.column === 3)?.cell,
    ).toEqual(original);

    controller.redo();
    expect(
      richCells(controller.getDocument()).find((item) => item.row === 2 && item.column === 3)?.cell,
    ).toMatchObject({
      input: { type: 'string', value: 'edited' },
      resourceId: 'resource-1',
      templateId: 'template-1',
      metadata: { untouched: true },
    });
  });

  it('restores a cleared rich cell through undo', () => {
    const controller = new SpreadsheetDocumentController(directDocument());
    const sheet = sheetId('sheet-1');
    const original = richCells(controller.getDocument()).find(
      (item) => item.row === 2 && item.column === 3,
    )!.cell;
    controller.dispatch(
      { type: 'clear-contents', selection: selection(sheet, 2, 3) },
      'context-menu',
    );
    controller.undo();

    expect(
      richCells(controller.getDocument()).find((item) => item.row === 2 && item.column === 3)?.cell,
    ).toEqual(original);
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

  it('rejects beforeNotify reentrant mutations without splitting document and projection truth', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = controller.getSnapshot();
    expect(() =>
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
          text: 'outer',
        },
        'ref',
        {
          beforeNotify: () => {
            controller.dispatch(
              {
                type: 'set-cell-text',
                address: { sheet: sheetId('sheet-1'), row: 1, column: 0 },
                text: 'inner',
              },
              'ref',
            );
          },
        },
      ),
    ).toThrow('Commit callback');

    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.getDocument().workbook.sheets[0]?.cells).toEqual([]);
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 0, column: 0 })).toBe('');
    const accepted = controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'accepted',
      },
      'ref',
    );
    expect(accepted.status).toBe('committed');
    if (accepted.status === 'committed') {
      expect(accepted.commit.change.id).toMatch(/^change-\d+-1$/);
    }
  });

  it('rejects beforeNotify lifecycle disposal without corrupting the active controller', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const before = controller.getSnapshot();
    expect(() =>
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
          text: 'outer',
        },
        'ref',
        { beforeNotify: () => controller.dispose() },
      ),
    ).toThrow('Commit callback');
    expect(controller.getSnapshot()).toEqual(before);
    expect(
      controller.dispatch(
        {
          type: 'set-cell-text',
          address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
          text: 'still-active',
        },
        'ref',
      ).status,
    ).toBe('committed');
  });

  it('rejects foreign and forged outer checkpoints before changing either runtime truth', () => {
    const first = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'first', sheetId: 'sheet-1' }),
    );
    const second = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'second', sheetId: 'sheet-2' }),
    );
    const foreign = first.checkpoint();
    const genuine = second.checkpoint();
    const before = second.getSnapshot();
    const forged = {
      ...genuine,
      document: first.getDocument(),
    };

    expect(() => second.restore(foreign)).toThrow('Checkpoint');
    expect(() => second.restore(forged)).toThrow('Checkpoint');
    expect(second.getSnapshot()).toEqual(before);
    expect(second.getDocument().id).toBe('second');
  });

  it('rolls back the legacy half when outer checkpoint restoration fails after it starts', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const checkpoint = controller.checkpoint();
    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'current',
      },
      'ref',
    );
    const before = controller.getSnapshot();
    const original = History.prototype.restore;
    let calls = 0;
    const restore = vi
      .spyOn(History.prototype, 'restore')
      .mockImplementation(function (this: History<unknown, unknown>, value) {
        calls += 1;
        if (calls === 2) throw new Error('schema restore failed');
        return original.call(this, value);
      });
    try {
      expect(() => controller.restore(checkpoint)).toThrow('schema restore failed');
    } finally {
      restore.mockRestore();
    }

    expect(controller.getSnapshot()).toEqual(before);
    expect(controller.getCellText({ sheet: sheetId('sheet-1'), row: 0, column: 0 })).toBe(
      'current',
    );
  });

  it('stores document patches in schema history instead of full document roots', () => {
    const controller = new SpreadsheetDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    controller.dispatch(
      {
        type: 'set-cell-text',
        address: { sheet: sheetId('sheet-1'), row: 0, column: 0 },
        text: 'committed',
      },
      'ref',
    );
    const entry = controller.checkpoint().documentHistory.undo[0]!;

    expect(entry.before).toMatchObject({ operations: expect.any(Array) });
    expect(entry.after).toMatchObject({ operations: expect.any(Array) });
    expect(entry.before).not.toHaveProperty('schemaVersion');
    expect(entry.after).not.toHaveProperty('schemaVersion');
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
