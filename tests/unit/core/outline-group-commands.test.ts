import { describe, expect, it, vi } from 'vitest';
import type { SheetId } from '../../../src/core';
import { projectDocumentToLegacy } from '../../../src/core/controller/runtime-projection';
import {
  parseSpreadsheetDocument,
  type GroupId,
  type SpreadsheetDocumentInput,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';

const sheet = 'sheet-1' as SheetId;

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
        id: 'insert-row',
        command: { type: 'insert-row', sheet, index: 3, count: 2 },
      }),
    ).toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.sheets[0]?.groups).toEqual([
      expect.objectContaining({ id: 'outer', start: 1, end: 10, level: 1 }),
      expect.objectContaining({ id: 'inner', start: 2, end: 6, level: 2 }),
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
});
