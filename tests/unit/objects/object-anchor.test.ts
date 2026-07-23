import { describe, expect, it } from 'vitest';
import type { DocumentSheetId, ObjectId, ResourceId } from '../../../src/document';
import {
  objectToDisplayCommands,
  resolveObjectAnchor,
  transformObjectAnchor,
} from '../../../src/objects';

describe('OBJ-01 shared object anchors', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;
  const geometry = {
    rowOffset: (row: number) => row * 20,
    columnOffset: (column: number) => column * 80,
  };

  it('resolves and transforms absolute, one-cell, and two-cell anchors', () => {
    expect(
      resolveObjectAnchor(
        {
          type: 'one-cell',
          cell: { sheetId, row: 2, column: 1 },
          offset: { x: 4, y: 5 },
          size: { width: 100, height: 40 },
        },
        geometry,
      ),
    ).toEqual({ x: 84, y: 45, width: 100, height: 40 });
    expect(
      transformObjectAnchor(
        {
          type: 'two-cell',
          from: { sheetId, row: 1, column: 1, offset: { x: 0, y: 0 } },
          to: { sheetId, row: 3, column: 2, offset: { x: 0, y: 0 } },
        },
        { type: 'insert-row', sheetId: 'sheet-1', index: 2, count: 2 },
      ),
    ).toMatchObject({
      from: { row: 1 },
      to: { row: 5 },
    });
  });

  it('emits safe ordered print commands and diagnoses missing resources', () => {
    expect(() =>
      objectToDisplayCommands(
        {
          id: 'logo' as ObjectId,
          kind: 'image',
          anchor: { type: 'absolute', rect: { x: 0, y: 0, width: 40, height: 20 } },
          zIndex: 2,
          locked: false,
          templateRepeat: 'shared',
          resourceId: 'missing' as ResourceId,
          accessibility: { name: 'Logo' },
        },
        { resources: {}, geometry },
      ),
    ).toThrowError(expect.objectContaining({ code: 'OBJECT_RESOURCE_MISSING' }));
    expect(
      objectToDisplayCommands(
        {
          id: 'notice' as ObjectId,
          kind: 'text-box',
          anchor: { type: 'absolute', rect: { x: 10, y: 10, width: 80, height: 20 } },
          zIndex: 1,
          locked: false,
          templateRepeat: 'shared',
          text: '<b>plain text</b>',
          style: { color: '#111111', fontFamily: 'Arial', fontSize: 12 },
          accessibility: { name: 'Notice' },
        },
        { resources: {}, geometry },
      ),
    ).toMatchObject([{ kind: 'text', text: '<b>plain text</b>' }]);
  });
});
