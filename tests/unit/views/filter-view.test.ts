import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import {
  applyDocumentFilterView,
  applyFilterView,
  createFilterViewSession,
} from '../../../src/views';
import { createSheetGridModel } from '../../../src/engine';
import { createDocumentController } from '../../../src/document-controller';
import { parseSpreadsheetDocument } from '../../../src/document';
import type { SheetId } from '../../../src/core';

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

  it('projects document input and formula values for renderer consumers', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'view-projection',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string', value: 'amount' } } },
              { row: 1, column: 0, cell: { input: { type: 'formula', source: '=10' } } },
              { row: 2, column: 0, cell: { input: { type: 'number', value: 3 } } },
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
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    expect(
      applyDocumentFilterView({
        document: parsed.document,
        formulaValues: new Map([['sheet-1!A2', { type: 'number', value: 10 }]]),
        view: {
          id: 'formula-view',
          name: 'Formula view',
          range: {
            sheetId,
            start: { row: 0, column: 0 },
            end: { row: 2, column: 0 },
          },
          sorts: [{ column: 0, direction: 'descending' }],
          filters: [{ column: 0, operator: 'greaterThan', value: 5 }],
          visibility: 'session',
        },
        locale: 'en-US',
        limits: { maxRows: 10 },
      }),
    ).toMatchObject({ rowOrder: [1, 2], hiddenRows: new Set([2]) });
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

  it('projects derived visibility and stable row order into grid geometry', () => {
    const model = createSheetGridModel(
      { rows: { len: 5 }, cols: { len: 2 } },
      {
        derivedRows: {
          start: 1,
          end: 3,
          rowOrder: [3, 1, 2],
          hiddenRows: new Set([2]),
        },
      },
    );
    expect(model.logicalRowAtVisualIndex(1)).toBe(3);
    expect(model.logicalRowAtVisualIndex(2)).toBe(1);
    expect(model.rowHeight(2)).toBe(0);
  });

  it('activates and deactivates saved views without changing document revision', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'view-controller',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [],
            merges: [],
            filterViews: [
              {
                id: 'saved',
                name: 'Saved',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 3, column: 1 },
                },
                sorts: [],
                filters: [],
                visibility: 'document',
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
    });
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    const controller = createDocumentController(parsed.document);
    const revision = controller.getSnapshot().revision;
    controller.activateFilterView('sheet-1' as SheetId, 'saved');
    expect(controller.getActiveFilterView('sheet-1' as SheetId)?.id).toBe('saved');
    expect(controller.getSnapshot().revision).toBe(revision);
    controller.deactivateFilterView('sheet-1' as SheetId);
    expect(controller.getActiveFilterView('sheet-1' as SheetId)).toBeUndefined();
    expect(controller.getSnapshot().revision).toBe(revision);
  });
});
