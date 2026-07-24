import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';
import type { SheetId } from '../../../src/core';
import { remapWorkbookCommand } from '../../../src/react/control/controlled-reconciler';

function fixture() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'phase-three-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [],
          merges: [],
          rowCount: 10,
          columnCount: 5,
          filterViews: [
            {
              id: 'view-1',
              name: 'Saved',
              range: {
                sheetId: 'sheet-1',
                start: { row: 0, column: 0 },
                end: { row: 3, column: 2 },
              },
              sorts: [{ column: 1, direction: 'ascending' }],
              filters: [{ column: 2, operator: 'contains', value: 'paid' }],
              visibility: 'document',
            },
          ],
          objects: [
            {
              id: 'object-1',
              kind: 'text-box',
              anchor: {
                type: 'one-cell',
                cell: { sheetId: 'sheet-1', row: 2, column: 1 },
                offset: { x: 4, y: 5 },
                size: { width: 100, height: 40 },
              },
              zIndex: 1,
              locked: false,
              templateRepeat: 'shared',
              text: 'Notice',
              style: { color: '#111111', fontFamily: 'Arial', fontSize: 12 },
              accessibility: { name: 'Notice' },
            },
          ],
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  } as unknown as SpreadsheetDocumentInput);
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return parsed.document;
}

describe('saved views and objects in Workbook 2.0', () => {
  it('round-trips persistent definitions without serializing session selection', () => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    expect(serialized.workbook.sheets[0]).toMatchObject({
      filterViews: [{ id: 'view-1', visibility: 'document' }],
      objects: [{ id: 'object-1', kind: 'text-box' }],
    });
    expect(serialized.workbook.sheets[0]).not.toHaveProperty('activeViewId');
    expect(parseSpreadsheetDocument(serialized)).toMatchObject({ ok: true });
  });

  it('transforms saved-view ranges and object anchors in one undoable F2 transaction', () => {
    const controller = createDocumentController(fixture());
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'insert-row',
        command: {
          type: 'insert-row',
          sheet: 'sheet-1' as SheetId,
          index: 1,
          count: 2,
        },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ range: { start: { row: 0 }, end: { row: 5 } } }],
      objects: [{ anchor: { cell: { row: 4 } } }],
    });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ range: { start: { row: 0 }, end: { row: 3 } } }],
      objects: [{ anchor: { cell: { row: 2 } } }],
    });
  });

  it.each([
    {
      name: 'absolute',
      anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
      command: { type: 'delete-row', sheet: 'sheet-1' as SheetId, index: 1, count: 2 },
      after: { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
    },
    {
      name: 'one-cell',
      anchor: {
        type: 'one-cell',
        cell: { sheetId: 'sheet-1', row: 2, column: 1 },
        offset: { x: 4, y: 5 },
        size: { width: 100, height: 40 },
      },
      command: { type: 'delete-row', sheet: 'sheet-1' as SheetId, index: 1, count: 2 },
      after: {
        type: 'one-cell',
        cell: { sheetId: 'sheet-1', row: 1, column: 1 },
        offset: { x: 4, y: 5 },
        size: { width: 100, height: 40 },
      },
    },
    {
      name: 'two-cell',
      anchor: {
        type: 'two-cell',
        from: {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          offset: { x: 0, y: 0 },
        },
        to: {
          sheetId: 'sheet-1',
          row: 4,
          column: 3,
          offset: { x: 10, y: 10 },
        },
      },
      command: { type: 'delete-column', sheet: 'sheet-1' as SheetId, index: 1, count: 2 },
      after: {
        type: 'two-cell',
        from: {
          sheetId: 'sheet-1',
          row: 1,
          column: 1,
          offset: { x: 0, y: 0 },
        },
        to: {
          sheetId: 'sheet-1',
          row: 4,
          column: 1,
          offset: { x: 0, y: 10 },
        },
      },
    },
  ])('preserves $name anchor transforms through undo and redo', ({ anchor, command, after }) => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    const sheet = serialized.workbook.sheets[0]!;
    sheet.objects = [{ ...sheet.objects![0]!, anchor }] as never;
    const parsed = parseSpreadsheetDocument(serialized);
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    const controller = createDocumentController(parsed.document);

    expect(
      controller.execute({
        schemaVersion: 1,
        id: `transform-${anchor.type}`,
        command: command as never,
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]!.objects[0]!.anchor).toEqual(after);
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]!.objects[0]!.anchor).toEqual(
      anchor,
    );
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]!.objects[0]!.anchor).toEqual(after);
  });

  it('normalizes collapsed two-cell offsets through delete, undo, and redo', () => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    const sheet = serialized.workbook.sheets[0]!;
    const anchor = {
      type: 'two-cell',
      from: {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        offset: { x: 10, y: 10 },
      },
      to: {
        sheetId: 'sheet-1',
        row: 2,
        column: 2,
        offset: { x: 0, y: 0 },
      },
    } as const;
    sheet.objects = [{ ...sheet.objects![0]!, anchor }] as never;
    const parsed = parseSpreadsheetDocument(serialized);
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    const controller = createDocumentController(parsed.document);

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'collapse-object-markers',
        command: {
          type: 'delete-row',
          sheet: 'sheet-1' as SheetId,
          index: 1,
          count: 2,
        },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]!.objects[0]!.anchor).toEqual({
      type: 'two-cell',
      from: {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        offset: { x: 10, y: 0 },
      },
      to: {
        sheetId: 'sheet-1',
        row: 1,
        column: 2,
        offset: { x: 0, y: 0 },
      },
    });
    expect(parseSpreadsheetDocument(controller.getSnapshot().document)).toMatchObject({ ok: true });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]!.objects[0]!.anchor).toEqual(
      anchor,
    );
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(parseSpreadsheetDocument(controller.getSnapshot().document)).toMatchObject({ ok: true });
  });

  it('rejects view/object limit overflow and dangling resources atomically', () => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    const viewOverflow = parseSpreadsheetDocument(serialized, { limits: { maxViews: 0 } });
    const objectOverflow = parseSpreadsheetDocument(serialized, { limits: { maxObjects: 0 } });
    expect(viewOverflow).toMatchObject({ ok: false });
    expect(viewOverflow).not.toHaveProperty('document');
    expect(objectOverflow).toMatchObject({ ok: false });
    expect(objectOverflow).not.toHaveProperty('document');

    const sheet = serialized.workbook.sheets[0];
    if (sheet === undefined) throw new Error('fixture sheet missing');
    sheet.objects = [
      {
        ...sheet.objects![0]!,
        kind: 'image',
        resourceId: 'missing',
      },
    ] as never;
    expect(parseSpreadsheetDocument(serialized)).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'DANGLING_REFERENCE' })],
    });
  });

  it('reports crossed two-cell marker geometry with a stable object diagnostic', () => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    const sheet = serialized.workbook.sheets[0]!;
    sheet.objects = [
      {
        ...sheet.objects![0]!,
        anchor: {
          type: 'two-cell',
          from: {
            sheetId: 'sheet-1',
            row: 4,
            column: 3,
            offset: { x: 10, y: 10 },
          },
          to: {
            sheetId: 'sheet-1',
            row: 2,
            column: 1,
            offset: { x: 0, y: 0 },
          },
        },
      },
    ] as never;

    expect(parseSpreadsheetDocument(serialized)).toMatchObject({
      ok: false,
      diagnostics: [
        expect.objectContaining({
          code: 'OBJECT_ANCHOR_INVALID',
          details: { path: '$.workbook.sheets[0].objects[0].anchor' },
        }),
      ],
    });
  });

  it.each([
    {
      from: {
        sheetId: 'sheet-1',
        row: 2,
        column: 1,
        offset: { x: 0, y: 10 },
      },
      to: {
        sheetId: 'sheet-1',
        row: 2,
        column: 2,
        offset: { x: 0, y: 0 },
      },
    },
    {
      from: {
        sheetId: 'sheet-1',
        row: 1,
        column: 2,
        offset: { x: 10, y: 0 },
      },
      to: {
        sheetId: 'sheet-1',
        row: 2,
        column: 2,
        offset: { x: 0, y: 0 },
      },
    },
  ])('rejects crossed same-row or same-column marker offsets', ({ from, to }) => {
    const serialized = JSON.parse(
      serializeSpreadsheetDocument(fixture()),
    ) as SpreadsheetDocumentInput;
    const sheet = serialized.workbook.sheets[0]!;
    sheet.objects = [
      {
        ...sheet.objects![0]!,
        anchor: { type: 'two-cell', from, to },
      },
    ] as never;

    expect(parseSpreadsheetDocument(serialized)).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'OBJECT_ANCHOR_INVALID' })],
    });
  });

  it('adds and removes views and objects only through undoable document commands', () => {
    const controller = createDocumentController(fixture());
    const sheet = 'sheet-1' as SheetId;
    expect(
      controller.transact({
        schemaVersion: 1,
        id: 'phase-three-definitions',
        baseRevision: 0,
        commands: [
          {
            schemaVersion: 1,
            id: 'set-view',
            command: {
              type: 'set-filter-view',
              sheet,
              view: {
                id: 'view-2',
                name: 'Second',
                range: {
                  sheetId: 'sheet-1' as never,
                  start: { row: 0, column: 0 },
                  end: { row: 2, column: 2 },
                },
                sorts: [],
                filters: [],
                visibility: 'document',
              },
            },
          },
          {
            schemaVersion: 1,
            id: 'set-object',
            command: {
              type: 'set-sheet-object',
              sheet,
              object: {
                id: 'object-2' as never,
                kind: 'text-box',
                anchor: {
                  type: 'absolute',
                  rect: { x: 1, y: 2, width: 30, height: 20 },
                },
                zIndex: 2,
                locked: false,
                templateRepeat: 'shared',
                text: 'Added',
                style: { color: '#000000', fontFamily: 'Arial', fontSize: 10 },
                accessibility: { name: 'Added' },
              },
            },
          },
        ],
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ id: 'view-1' }, { id: 'view-2' }],
      objects: [{ id: 'object-1' }, { id: 'object-2' }],
    });
    expect(
      controller.transact({
        schemaVersion: 1,
        id: 'phase-three-definition-removal',
        baseRevision: 1,
        commands: [
          {
            schemaVersion: 1,
            id: 'remove-view',
            command: { type: 'remove-filter-view', sheet, viewId: 'view-2' },
          },
          {
            schemaVersion: 1,
            id: 'remove-object',
            command: { type: 'remove-sheet-object', sheet, objectId: 'object-2' },
          },
        ],
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ id: 'view-1' }],
      objects: [{ id: 'object-1' }],
    });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ id: 'view-1' }, { id: 'view-2' }],
      objects: [{ id: 'object-1' }, { id: 'object-2' }],
    });
    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]).toMatchObject({
      filterViews: [{ id: 'view-1' }],
      objects: [{ id: 'object-1' }],
    });
  });

  it('remaps command ownership and nested anchors for controlled replay', () => {
    const original = 'sheet-1' as SheetId;
    const replayed = 'sheet-replayed' as SheetId;
    const mapping = new Map([[original, replayed]]);
    expect(
      remapWorkbookCommand(
        {
          type: 'set-filter-view',
          sheet: original,
          view: fixture().workbook.sheets[0]!.filterViews[0]!,
        },
        mapping,
      ),
    ).toMatchObject({
      sheet: replayed,
      view: { range: { sheetId: replayed } },
    });
    expect(
      remapWorkbookCommand(
        {
          type: 'set-sheet-object',
          sheet: original,
          object: fixture().workbook.sheets[0]!.objects[0]!,
        },
        mapping,
      ),
    ).toMatchObject({
      sheet: replayed,
      object: { anchor: { cell: { sheetId: replayed } } },
    });
  });
});
