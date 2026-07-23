import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';
import type { SheetId } from '../../../src/core';

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
});
