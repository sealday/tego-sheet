import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';

function fixture(groups: readonly unknown[]): SpreadsheetDocumentInput {
  return {
    schemaVersion: 2,
    id: 'outline-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [],
          merges: [],
          rowCount: 10,
          columnCount: 6,
          groups,
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  } as unknown as SpreadsheetDocumentInput;
}

describe('Workbook 2.0 outline groups', () => {
  it('normalizes nested levels and round-trips groups in canonical order', () => {
    const parsed = parseSpreadsheetDocument(
      fixture([
        {
          id: 'row-inner',
          axis: 'row',
          start: 2,
          end: 4,
          level: 99,
          collapsed: true,
        },
        {
          id: 'column',
          axis: 'column',
          start: 1,
          end: 3,
          level: 7,
          collapsed: false,
        },
        {
          id: 'row-outer',
          axis: 'row',
          start: 1,
          end: 8,
          level: 4,
          collapsed: false,
        },
      ]),
    );

    expect(parsed).toMatchObject({
      ok: true,
      document: {
        workbook: {
          sheets: [
            {
              groups: [
                { id: 'row-outer', axis: 'row', start: 1, end: 8, level: 1 },
                { id: 'row-inner', axis: 'row', start: 2, end: 4, level: 2 },
                { id: 'column', axis: 'column', start: 1, end: 3, level: 1 },
              ],
            },
          ],
        },
      },
    });
    if (!parsed.ok) throw new Error('outline fixture must parse');
    expect(
      parseSpreadsheetDocument(JSON.parse(serializeSpreadsheetDocument(parsed.document))),
    ).toMatchObject({ ok: true, document: parsed.document });
  });

  it.each([
    {
      name: 'duplicate IDs',
      groups: [
        { id: 'same', axis: 'row', start: 0, end: 1, level: 1, collapsed: false },
        { id: 'same', axis: 'column', start: 0, end: 1, level: 1, collapsed: false },
      ],
      code: 'DUPLICATE_ID',
    },
    {
      name: 'crossing ranges',
      groups: [
        { id: 'left', axis: 'row', start: 0, end: 4, level: 1, collapsed: false },
        { id: 'right', axis: 'row', start: 3, end: 6, level: 1, collapsed: false },
      ],
      code: 'GROUP_LIMIT_EXCEEDED',
    },
    {
      name: 'out-of-bounds ranges',
      groups: [{ id: 'outside', axis: 'column', start: 2, end: 6, level: 1, collapsed: false }],
      code: 'GROUP_LIMIT_EXCEEDED',
    },
  ])('rejects $name', ({ groups, code }) => {
    expect(parseSpreadsheetDocument(fixture(groups))).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code })],
    });
  });

  it('enforces the configurable group resource limit before decoding entries', () => {
    expect(
      parseSpreadsheetDocument(
        fixture([
          { id: 'one', axis: 'row', start: 0, end: 1, level: 1, collapsed: false },
          { id: 'two', axis: 'row', start: 3, end: 4, level: 1, collapsed: false },
        ]),
        { limits: { maxGroups: 1 } },
      ),
    ).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'GROUP_LIMIT_EXCEEDED' })],
    });
  });
});
