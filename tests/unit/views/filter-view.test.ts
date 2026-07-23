import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import { applyFilterView, createFilterViewSession } from '../../../src/views';

describe('VIEW-01 derived filter views', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;
  const rows = [
    [
      { type: 'string' as const, value: 'name' },
      { type: 'string' as const, value: 'amount' },
    ],
    [
      { type: 'string' as const, value: 'A' },
      { type: 'number' as const, value: 10 },
    ],
    [
      { type: 'string' as const, value: 'B' },
      { type: 'number' as const, value: 3 },
    ],
    [
      { type: 'string' as const, value: 'C' },
      { type: 'number' as const, value: 10 },
    ],
  ];

  it('derives visibility without persisting hidden row truth', () => {
    const result = applyFilterView({
      view: {
        id: 'high-value',
        name: 'High value',
        range: {
          sheetId,
          start: { row: 0, column: 0 },
          end: { row: 3, column: 1 },
        },
        sorts: [],
        filters: [{ column: 1, operator: 'greaterThanOrEqual', value: 10 }],
        visibility: 'document',
      },
      rows,
      locale: 'en-US',
      limits: { maxRows: 100 },
    });
    expect([...result.hiddenRows]).toEqual([2]);
    expect(rows).toHaveLength(4);
  });

  it('derives a stable multi-column row order without moving source rows', () => {
    const result = applyFilterView({
      view: {
        id: 'sorted',
        name: 'Sorted',
        range: {
          sheetId,
          start: { row: 0, column: 0 },
          end: { row: 3, column: 1 },
        },
        sorts: [
          { column: 1, direction: 'descending' },
          { column: 0, direction: 'descending' },
        ],
        filters: [],
        visibility: 'document',
      },
      rows,
      locale: 'en-US',
      limits: { maxRows: 100 },
    });
    expect(result.rowOrder).toEqual([3, 1, 2]);
    expect(rows[1]?.[0]).toMatchObject({ value: 'A' });
  });

  it('keeps session view selection outside document change state', () => {
    const session = createFilterViewSession();
    const before = session.revision;
    session.select({
      id: 'personal',
      name: 'Personal',
      range: {
        sheetId,
        start: { row: 0, column: 0 },
        end: { row: 3, column: 1 },
      },
      sorts: [],
      filters: [],
      visibility: 'session',
    });
    expect(session.selected?.id).toBe('personal');
    expect(session.revision).toBe(before + 1);
    expect(session.documentRevision).toBe(0);
  });
});
