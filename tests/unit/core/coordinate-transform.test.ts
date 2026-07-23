import { describe, expect, it, vi } from 'vitest';
import {
  CoordinateTransform,
  transformDocumentCoordinates,
  transformSheetCoordinates,
} from '../../../src/core/coordinates/coordinate-transform';
import {
  parseSpreadsheetDocument,
  type SheetInput,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { sheetId } from '../../../src/core';
import { SpreadsheetDocumentController } from '../../../src/core/controller/spreadsheet-document-controller';

function sheet(): SheetInput {
  return {
    id: 'sheet-1',
    name: 'Data',
    rowCount: 20,
    columnCount: 10,
    cells: [
      {
        row: 1,
        column: 0,
        cell: { input: { type: 'formula', source: '=A3+$B$4' }, validationId: 'validation-1' },
      },
      { row: 2, column: 1, cell: { input: { type: 'string', value: 'deleted' } } },
      {
        row: 3,
        column: 3,
        cell: { input: { type: 'formula', source: '=B1+D1' } },
      },
    ],
    rows: [{ index: 2, height: 30 }],
    columns: [
      { index: 1, width: 70 },
      { index: 3, width: 90 },
    ],
    merges: [{ start: { row: 1, column: 0 }, end: { row: 3, column: 3 } }],
    freeze: { row: 3, column: 3 },
    filter: {
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 4 } },
      filters: [
        { column: 1, operator: 'all', values: [] },
        { column: 3, operator: 'in', values: ['kept'] },
      ],
      sort: { column: 2, direction: 'asc' },
    },
  };
}

describe('CoordinateTransform', () => {
  it('rejects unsafe transforms and impossible structural counts', () => {
    expect(() => CoordinateTransform.insert('row', -1, 1)).toThrow();
    expect(() => CoordinateTransform.insert('row', 0, 0)).toThrow();
    expect(() => CoordinateTransform.delete('column', 0, Number.POSITIVE_INFINITY)).toThrow();
    expect(() =>
      transformSheetCoordinates(
        { ...sheet(), rowCount: 1 },
        CoordinateTransform.delete('row', 0, 2),
      ),
    ).toThrow();
  });

  it('matches sheet names without consulting the host locale', () => {
    const localeFold = vi.spyOn(String.prototype, 'toLocaleLowerCase').mockImplementation(() => {
      throw new Error('locale-sensitive fold used');
    });
    try {
      expect(
        CoordinateTransform.insert('row', 0).formula('=I!A1', {
          targetSheetName: 'I',
          transformUnqualified: false,
        }),
      ).toBe('=I!A2');
    } finally {
      localeFold.mockRestore();
    }
  });

  it('keeps formula function names intact while transforming their arguments', () => {
    const transform = CoordinateTransform.insert('row', 1, 2);

    expect(transform.formula('=LOG10(A2)+DAYS360(A2,B3)+SUM(A2:B3)')).toBe(
      '=LOG10(A4)+DAYS360(A4,B5)+SUM(A4:B5)',
    );
  });

  it('shrinks or drops formula ranges as one coordinate range during deletion', () => {
    const transform = CoordinateTransform.delete('row', 1, 2);

    expect(transform.formula('=A1:A4+B2:B3+C4:C5')).toBe('=A1:A2+#REF!+C2:C3');
    expect(
      transform.formula('=Data!A1:A4', {
        targetSheetName: 'Data',
        transformUnqualified: false,
      }),
    ).toBe('=Data!A1:A2');
  });

  it('uses one insert transform for cells, formulas, ranges, layout, freeze, and filters', () => {
    const transform = CoordinateTransform.insert('row', 2, 2);
    const next = transformSheetCoordinates(sheet(), transform);

    expect(next.cells.map(({ row, column }) => [row, column])).toEqual([
      [1, 0],
      [4, 1],
      [5, 3],
    ]);
    expect(next.cells[0]?.cell.input).toEqual({ type: 'formula', source: '=A5+$B$6' });
    expect(next.cells[0]?.cell.validationId).toBe('validation-1');
    expect(next.rows).toEqual([{ index: 4, height: 30 }]);
    expect(next.merges).toEqual([{ start: { row: 1, column: 0 }, end: { row: 5, column: 3 } }]);
    expect(next.freeze).toEqual({ row: 5, column: 3 });
    expect(next.filter?.range).toEqual({
      start: { row: 0, column: 0 },
      end: { row: 7, column: 4 },
    });
  });

  it('uses one delete transform to drop covered points and shrink every surviving range', () => {
    const transform = CoordinateTransform.delete('column', 1, 2);
    const next = transformSheetCoordinates(sheet(), transform);

    expect(next.cells.map(({ row, column }) => [row, column])).toEqual([
      [1, 0],
      [3, 1],
    ]);
    expect(next.cells[0]?.cell.input).toEqual({ type: 'formula', source: '=A3+#REF!' });
    expect(next.cells[1]?.cell.input).toEqual({ type: 'formula', source: '=#REF!+B1' });
    expect(next.columns).toEqual([{ index: 1, width: 90 }]);
    expect(next.merges).toEqual([{ start: { row: 1, column: 0 }, end: { row: 3, column: 1 } }]);
    expect(next.freeze).toEqual({ row: 3, column: 1 });
    expect(next.filter).toEqual({
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 2 } },
      filters: [{ column: 1, operator: 'in', values: ['kept'] }],
      sort: null,
    });
  });

  it('does not materialize absent optional filter fields during a column transform', () => {
    const source = sheet();
    source.filter = {
      range: source.filter?.range,
      filters: [],
    };
    const next = transformSheetCoordinates(source, CoordinateTransform.insert('column', 1));

    expect(next.filter).not.toHaveProperty('sort');
  });

  it('transforms template and print ranges through the same range primitive', () => {
    const transform = CoordinateTransform.delete('row', 2, 2);
    expect(
      transform.range({
        start: { row: 1, column: 1 },
        end: { row: 5, column: 4 },
      }),
    ).toEqual({
      start: { row: 1, column: 1 },
      end: { row: 3, column: 4 },
    });
    expect(
      transform.range({
        start: { row: 2, column: 1 },
        end: { row: 3, column: 4 },
      }),
    ).toBeNull();
  });

  it('transforms local and explicit cross-sheet references without touching other sheet qualifiers', () => {
    const input: SpreadsheetDocumentInput = {
      schemaVersion: 2,
      id: 'document-1',
      workbook: {
        sheets: [
          {
            ...sheet(),
            name: 'Data 2026',
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: {
                    type: 'formula',
                    source: `=A3+'Data 2026'!$A$3:Other!B4+"A3"`,
                  },
                },
              },
            ],
          },
          {
            id: 'sheet-2',
            name: 'Other',
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: {
                    type: 'formula',
                    source: `='Data 2026'!A3:B4+Other!A3`,
                  },
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
      templates: [
        {
          id: 'template-1',
          name: 'Print',
          sheetId: 'sheet-1',
          range: {
            sheetId: 'sheet-1',
            start: { row: 1, column: 0 },
            end: { row: 4, column: 2 },
          },
          printProfile: {
            paperSize: 'A4',
            orientation: 'portrait',
            margins: { top: 1, right: 1, bottom: 1, left: 1 },
          },
        },
      ],
      resources: { items: [] },
      extensions: {},
    };

    const next = transformDocumentCoordinates(
      input,
      'sheet-1',
      CoordinateTransform.insert('row', 2, 2),
    );

    expect(next.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
      type: 'formula',
      source: `=A5+'Data 2026'!$A$5:Other!B4+"A3"`,
    });
    expect(next.workbook.sheets[1]?.cells[0]?.cell.input).toEqual({
      type: 'formula',
      source: `='Data 2026'!A5:B6+Other!A3`,
    });
    expect(next.templates[0]?.range).toEqual({
      sheetId: 'sheet-1',
      start: { row: 1, column: 0 },
      end: { row: 6, column: 2 },
    });

    const parsed = parseSpreadsheetDocument(input);
    if (!parsed.ok) throw new Error('Cross-sheet coordinate fixture must be valid');
    const controller = new SpreadsheetDocumentController(parsed.document);
    controller.dispatch(
      { type: 'insert-row', sheet: sheetId('sheet-1'), index: 2, count: 2 },
      'ref',
    );
    const committed = controller.getDocument();
    expect(committed.workbook.sheets[0]?.cells[0]?.cell.input).toEqual(
      next.workbook.sheets[0]?.cells[0]?.cell.input,
    );
    expect(committed.workbook.sheets[1]?.cells[0]?.cell.input).toEqual(
      next.workbook.sheets[1]?.cells[0]?.cell.input,
    );
    expect(committed.templates[0]?.range).toEqual(next.templates[0]?.range);
  });
});
