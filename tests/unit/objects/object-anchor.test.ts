import { describe, expect, it } from 'vitest';
import type { DocumentSheetId, ObjectId, ResourceId } from '../../../src/document';
import type { ObjectAnchor, ObjectCoordinateTransform } from '../../../src/objects';
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

  it.each<readonly [string, ObjectAnchor, ObjectCoordinateTransform, ObjectAnchor]>([
    [
      'keeps absolute anchors fixed across row insertion',
      { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
      { type: 'insert-row', sheetId, index: 1, count: 2 },
      { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
    ],
    [
      'keeps absolute anchors fixed across column deletion',
      { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
      { type: 'delete-column', sheetId, index: 1, count: 2 },
      { type: 'absolute', rect: { x: 10, y: 20, width: 30, height: 40 } },
    ],
    [
      'moves one-cell anchors after row insertion',
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
      { type: 'insert-row', sheetId, index: 2, count: 2 },
      {
        type: 'one-cell',
        cell: { sheetId, row: 4, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
    ],
    [
      'clamps one-cell anchors deleted at a row boundary',
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
      { type: 'delete-row', sheetId, index: 1, count: 2 },
      {
        type: 'one-cell',
        cell: { sheetId, row: 1, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
    ],
    [
      'moves one-cell anchors after column insertion',
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
      { type: 'insert-column', sheetId, index: 1, count: 2 },
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 4 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
    ],
    [
      'clamps one-cell anchors deleted at a column boundary',
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
      { type: 'delete-column', sheetId, index: 2, count: 1 },
      {
        type: 'one-cell',
        cell: { sheetId, row: 2, column: 2 },
        offset: { x: 1, y: 2 },
        size: { width: 30, height: 40 },
      },
    ],
    [
      'expands two-cell anchors across row insertion',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 3, y: 4 } },
      },
      { type: 'insert-row', sheetId, index: 2, count: 2 },
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 6, column: 4, offset: { x: 3, y: 4 } },
      },
    ],
    [
      'clamps two-cell markers across row deletion',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 3, y: 4 } },
      },
      { type: 'delete-row', sheetId, index: 0, count: 2 },
      {
        type: 'two-cell',
        from: { sheetId, row: 0, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 2, column: 4, offset: { x: 3, y: 4 } },
      },
    ],
    [
      'expands two-cell anchors across column insertion',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 3, y: 4 } },
      },
      { type: 'insert-column', sheetId, index: 3, count: 1 },
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 5, offset: { x: 3, y: 4 } },
      },
    ],
    [
      'collapses two-cell markers deleted at column boundaries',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 3, y: 4 } },
      },
      { type: 'delete-column', sheetId, index: 1, count: 4 },
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 0, y: 2 } },
        to: { sheetId, row: 4, column: 1, offset: { x: 0, y: 4 } },
      },
    ],
    [
      'normalizes offsets when row deletion collapses both markers',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 10 } },
        to: { sheetId, row: 2, column: 4, offset: { x: 3, y: 0 } },
      },
      { type: 'delete-row', sheetId, index: 1, count: 2 },
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 1, offset: { x: 1, y: 0 } },
        to: { sheetId, row: 1, column: 4, offset: { x: 3, y: 0 } },
      },
    ],
    [
      'preserves offsets when deletion moves an already same-row anchor',
      {
        type: 'two-cell',
        from: { sheetId, row: 5, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 5, column: 4, offset: { x: 3, y: 12 } },
      },
      { type: 'delete-row', sheetId, index: 1, count: 1 },
      {
        type: 'two-cell',
        from: { sheetId, row: 4, column: 1, offset: { x: 1, y: 2 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 3, y: 12 } },
      },
    ],
    [
      'preserves offsets when deletion moves an already same-column anchor',
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 5, offset: { x: 2, y: 1 } },
        to: { sheetId, row: 4, column: 5, offset: { x: 12, y: 3 } },
      },
      { type: 'delete-column', sheetId, index: 1, count: 1 },
      {
        type: 'two-cell',
        from: { sheetId, row: 1, column: 4, offset: { x: 2, y: 1 } },
        to: { sheetId, row: 4, column: 4, offset: { x: 12, y: 3 } },
      },
    ],
  ])('%s', (_name, anchor, operation, expected) => {
    expect(transformObjectAnchor(anchor, operation)).toEqual(expected);
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
