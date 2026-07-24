import { describe, expect, it, vi } from 'vitest';
import type { SheetId } from '../../../src/core';
import {
  projectDocumentToLegacy,
  projectLegacyToDocument,
} from '../../../src/core/controller/runtime-projection';
import {
  parseSpreadsheetDocument,
  type GroupId,
  type SpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';

const sheet = 'sheet-1' as SheetId;

function legacyHidden(
  layouts: Record<string, unknown> | undefined,
  index: number,
): boolean | undefined {
  return (layouts?.[String(index)] as { readonly hide?: boolean } | undefined)?.hide;
}

function controllerFixture() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'outline-command-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [],
          merges: [],
          rowCount: 12,
          columnCount: 8,
          rows: [{ index: 9, hidden: true }],
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  } as SpreadsheetDocumentInput);
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return createDocumentController(parsed.document);
}

function groupCommand(id: string, start: number, end: number, axis: 'row' | 'column' = 'row') {
  return {
    schemaVersion: 1 as const,
    id: `group-${id}`,
    command: {
      type: 'group' as const,
      sheet,
      group: { id: id as GroupId, axis, start, end, collapsed: false },
    },
  };
}

describe('versioned outline group commands', () => {
  it('groups, toggles, ungroups, and preserves history and permission boundaries', () => {
    const controller = controllerFixture();
    const denied = vi.fn(() => false);
    expect(
      controller.execute(groupCommand('outer', 1, 6), { permissionGate: denied }),
    ).toMatchObject({ status: 'rejected', code: 'COMMAND_NOT_ALLOWED' });
    expect(denied).toHaveBeenCalledTimes(1);
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([]);

    expect(controller.execute(groupCommand('outer', 1, 6))).toMatchObject({
      status: 'committed',
    });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'toggle-outer',
        command: { type: 'toggle-group', sheet, id: 'outer' as GroupId },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([
      {
        id: 'outer',
        axis: 'row',
        start: 1,
        end: 6,
        level: 1,
        collapsed: true,
      },
    ]);

    const document = controller.getSnapshot().document;
    const runtime = projectDocumentToLegacy(document)[0]!;
    expect(
      Array.from({ length: 12 }, (_, index) => index).filter(
        (index) =>
          ((runtime.rows ?? {})[String(index)] as { readonly hide?: boolean } | undefined)?.hide ===
          true,
      ),
    ).toEqual([1, 2, 3, 4, 5, 6, 9]);
    expect(document.workbook.sheets[0]?.rows).toEqual([{ index: 9, hidden: true }]);

    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups[0]?.collapsed).toBe(false);
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups[0]?.collapsed).toBe(true);

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'ungroup-outer',
        command: { type: 'ungroup', sheet, id: 'outer' as GroupId },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([]);
  });

  it('transforms endpoints, deletes empty groups, and recomputes nested levels', () => {
    const controller = controllerFixture();
    expect(controller.execute(groupCommand('outer', 1, 8))).toMatchObject({ status: 'committed' });
    expect(controller.execute(groupCommand('inner', 2, 4))).toMatchObject({ status: 'committed' });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'collapse-outer',
        command: { type: 'toggle-group', sheet, id: 'outer' as GroupId },
      }),
    ).toMatchObject({ status: 'committed' });

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'insert-row',
        command: { type: 'insert-row', sheet, index: 3, count: 2 },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([
      expect.objectContaining({ id: 'outer', start: 1, end: 10, level: 1 }),
      expect.objectContaining({ id: 'inner', start: 2, end: 6, level: 2 }),
    ]);
    expect(controller.getSnapshot().document.workbook.sheets[0]?.rows).toEqual([
      { index: 11, hidden: true },
    ]);

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'delete-inner',
        command: { type: 'delete-row', sheet, index: 2, count: 5 },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([
      expect.objectContaining({ id: 'outer', start: 1, end: 5, level: 1 }),
    ]);
  });

  it('rejects duplicate IDs, illegal crossings, missing IDs, and excessive nesting stably', () => {
    const controller = controllerFixture();
    expect(controller.execute(groupCommand('one', 0, 4))).toMatchObject({ status: 'committed' });
    expect(controller.execute(groupCommand('one', 6, 7))).toMatchObject({
      status: 'rejected',
      code: 'GROUP_LIMIT_EXCEEDED',
    });
    expect(controller.execute(groupCommand('cross', 3, 6))).toMatchObject({
      status: 'rejected',
      code: 'GROUP_LIMIT_EXCEEDED',
    });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'missing',
        command: { type: 'toggle-group', sheet, id: 'missing' as GroupId },
      }),
    ).toMatchObject({ status: 'rejected', code: 'GROUP_LIMIT_EXCEEDED' });
  });

  it('strips derived hide when resizing or editing inside collapsed groups', () => {
    const controller = controllerFixture();
    expect(controller.execute(groupCommand('rows', 1, 3))).toMatchObject({ status: 'committed' });
    expect(controller.execute(groupCommand('columns', 1, 3, 'column'))).toMatchObject({
      status: 'committed',
    });
    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `collapse-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'resize-row',
        command: { type: 'set-row-height', sheet, row: 2, height: 44 },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'resize-column',
        command: { type: 'set-column-width', sheet, column: 2, width: 88 },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'edit-group-cell',
        command: {
          type: 'set-cell-input',
          address: { sheet, row: 2, column: 2 },
          input: { type: 'string', value: 'edited' },
        },
      }),
    ).toMatchObject({ status: 'committed' });

    expect(controller.getSnapshot().document.workbook.sheets[0]?.rows).toEqual([
      { index: 2, height: 44 },
      { index: 9, hidden: true },
    ]);
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toEqual([
      { index: 2, width: 88 },
    ]);

    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `expand-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }
    const runtime = projectDocumentToLegacy(controller.getSnapshot().document)[0]!;
    expect((runtime.rows?.['2'] as { readonly hide?: boolean } | undefined)?.hide).not.toBe(true);
    expect((runtime.cols?.['2'] as { readonly hide?: boolean } | undefined)?.hide).not.toBe(true);
  });

  it('persists explicit row and column visibility inside collapsed groups across history', () => {
    const controller = controllerFixture();
    expect(controller.execute(groupCommand('rows', 1, 3))).toMatchObject({ status: 'committed' });
    expect(controller.execute(groupCommand('columns', 1, 3, 'column'))).toMatchObject({
      status: 'committed',
    });
    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `collapse-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'explicit-hide-row',
        command: { type: 'set-row-hidden', sheet, row: 2, hidden: true },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'explicit-hide-column',
        command: { type: 'set-column-hidden', sheet, column: 2, hidden: true },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.rows).toContainEqual({
      index: 2,
      hidden: true,
    });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toContainEqual({
      index: 2,
      hidden: true,
    });

    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).not.toContainEqual({
      index: 2,
      hidden: true,
    });
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toContainEqual({
      index: 2,
      hidden: true,
    });

    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `expand-explicit-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }
    const explicitlyHiddenRuntime = projectDocumentToLegacy(controller.getSnapshot().document)[0]!;
    expect(legacyHidden(explicitlyHiddenRuntime.rows, 2)).toBe(true);
    expect(legacyHidden(explicitlyHiddenRuntime.cols, 2)).toBe(true);
    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `recollapse-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }

    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'explicit-show-row',
        command: { type: 'set-row-hidden', sheet, row: 2, hidden: false },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(
      controller.execute({
        schemaVersion: 1,
        id: 'explicit-show-column',
        command: { type: 'set-column-hidden', sheet, column: 2, hidden: false },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.rows).toContainEqual({
      index: 2,
      hidden: false,
    });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toContainEqual({
      index: 2,
      hidden: false,
    });
    const collapsedRuntime = projectDocumentToLegacy(controller.getSnapshot().document)[0]!;
    expect(legacyHidden(collapsedRuntime.rows, 2)).toBe(true);
    expect(legacyHidden(collapsedRuntime.cols, 2)).toBe(true);

    expect(controller.undo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toContainEqual({
      index: 2,
      hidden: true,
    });
    expect(controller.redo()).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.columns).toContainEqual({
      index: 2,
      hidden: false,
    });

    for (const id of ['rows', 'columns']) {
      expect(
        controller.execute({
          schemaVersion: 1,
          id: `expand-${id}`,
          command: { type: 'toggle-group', sheet, id: id as GroupId },
        }),
      ).toMatchObject({ status: 'committed' });
    }
    const expandedRuntime = projectDocumentToLegacy(controller.getSnapshot().document)[0]!;
    expect(legacyHidden(expandedRuntime.rows, 2)).toBe(false);
    expect(legacyHidden(expandedRuntime.cols, 2)).toBe(false);
  });

  it('merges 10k collapsed groups with 100k layouts without per-layout group scans', () => {
    const groups = Array.from({ length: 10_000 }, (_, index) => ({
      id: `group-${index}` as GroupId,
      axis: 'row' as const,
      start: index * 10,
      end: index * 10,
      level: 1,
      collapsed: true,
    }));
    Object.defineProperty(groups, 'some', {
      value: () => {
        throw new Error('collapsed groups must be indexed before layout merge');
      },
    });
    const largeDocument = {
      schemaVersion: 2,
      id: 'large-outline-projection',
      workbook: {
        sheets: [
          {
            id: sheet,
            name: 'Sheet 1',
            visibility: 'visible',
            cells: [],
            merges: [],
            rowCount: 100_000,
            columnCount: 1,
            rows: Array.from({ length: 100_000 }, (_, index) => ({
              index,
              height: 20 + (index % 3),
            })),
            columns: [],
            groups,
            conditionalFormatting: [],
            filterViews: [],
            objects: [],
          },
        ],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' as const },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    } as unknown as SpreadsheetDocument;
    const legacy = projectDocumentToLegacy(largeDocument);
    const started = performance.now();
    const projected = projectLegacyToDocument(legacy, legacy, largeDocument, [sheet]);

    expect(projected.workbook.sheets[0]?.rows).toHaveLength(100_000);
    expect(performance.now() - started).toBeLessThan(10_000);
  }, 30_000);
});
