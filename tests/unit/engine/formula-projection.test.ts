import { describe, expect, it } from 'vitest';
import { sortRows } from '../../../src/core/operations/sort';
import { resolveCellPresentation } from '../../../src/engine/canvas/cell-painter';
import type { SheetData } from '../../../src/core/types/workbook';

const locale = { id: 'en-US', messages: {} };

describe('schema2 formula projection consumers', () => {
  it('renders a projected typed formula cache without invoking the legacy evaluator', () => {
    const sheet: SheetData = {
      rows: {
        0: { cells: { 0: { text: '=1/0', value: 42 } } },
      },
    };
    expect(resolveCellPresentation(sheet, { row: 0, column: 0 }, false).text).toBe('42');
  });

  it('sorts projected formula cells by their typed cache', () => {
    const sheet: SheetData = {
      rows: {
        1: { cells: { 0: { text: '=1/0', value: 2 } } },
        2: { cells: { 0: { text: '=1/0', value: 1 } } },
      },
    };
    expect(
      sortRows(sheet, 0, 'asc', locale, {
        start: { row: 0, column: 0 },
        end: { row: 2, column: 0 },
      }),
    ).toEqual([2, 1]);
  });
});
