import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet, type TegoSheetHandle } from '../../../src';
import type { ToolbarRenderProps } from '../../../src';
import type { SheetTabsRenderProps } from '../../../src';
import {
  createPermissionSnapshot,
  createPermissionStore,
} from '../../../src/integrations/permission';
import { createCanvasHarness } from '../../helpers/canvas-harness';
import { testDocument } from '../../helpers/workbook-builders';
import type { SpreadsheetTemplate } from '../../../src/template';

beforeEach(() => {
  const context = createCanvasHarness().canvas.getContext('2d');
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => context);
  vi.stubGlobal(
    'requestAnimationFrame',
    vi.fn(() => 1),
  );
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

it('uses one permission store for ref and toolbar commands with live default-deny updates', async () => {
  const ref = createRef<TegoSheetHandle>();
  const store = createPermissionStore();
  const errors = vi.fn();
  let toolbar: ToolbarRenderProps | undefined;
  let tabs: SheetTabsRenderProps | undefined;
  const source = testDocument([{ name: 'A' }]);
  const sheet = source.workbook.sheets[0]!.id as never;
  store.replace(
    createPermissionSnapshot({
      revision: 'permission-view',
      actorId: 'actor-1',
      grants: [
        {
          action: 'document:view',
          target: { type: 'document', documentId: source.id },
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId: sheet },
        },
      ],
    }),
  );
  render(
    <TegoSheet
      ref={ref}
      defaultDocument={source}
      permissionStore={store}
      toolbar={(props) => {
        toolbar = props;
        return null;
      }}
      sheetTabs={(props) => {
        tabs = props;
        return null;
      }}
      onError={errors}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());

  expect(toolbar!.disabledActions.has('set-style')).toBe(true);
  expect(() => ref.current!.setCellText({ sheet, row: 0, column: 0 }, 'denied')).toThrow(
    /permission/u,
  );
  expect(ref.current!.getCell({ sheet, row: 0, column: 0 })).toBeNull();

  act(() => {
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-document-only',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:edit',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'document:view',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'sheet:view',
            target: { type: 'sheet', sheetId: sheet },
          },
        ],
      }),
    );
  });
  await waitFor(() => expect(toolbar!.readOnly).toBe(false));
  expect(toolbar!.disabledActions.has('set-style')).toBe(true);
  expect(tabs!.readOnly).toBe(true);

  act(() => {
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-1',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:edit',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'document:view',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'sheet:view',
            target: { type: 'sheet', sheetId: sheet },
          },
          {
            action: 'range:edit',
            target: {
              type: 'range',
              range: {
                sheetId: sheet,
                start: { row: 0, column: 0 },
                end: { row: 0, column: 0 },
              },
            },
          },
          {
            action: 'sheet:edit',
            target: { type: 'sheet', sheetId: sheet },
          },
        ],
      }),
    );
  });
  await waitFor(() => expect(tabs!.readOnly).toBe(false));
  act(() => ref.current!.setCellText({ sheet, row: 0, column: 0 }, 'allowed'));
  expect(ref.current!.getCell({ sheet, row: 0, column: 0 })?.text).toBe('allowed');
});

it('defaults template changes to denied and requires template:bind permission', async () => {
  const source = testDocument([{ name: 'A' }]);
  const sheetId = source.workbook.sheets[0]!.id as never;
  const template: SpreadsheetTemplate = {
    id: 'template-1' as never,
    name: 'Invoice',
    bindings: [],
    printProfiles: [],
  };
  const store = createPermissionStore();
  store.replace(
    createPermissionSnapshot({
      revision: 'permission-view',
      actorId: 'actor-1',
      grants: [
        {
          action: 'document:view',
          target: { type: 'document', documentId: source.id },
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId },
        },
      ],
    }),
  );
  const onTemplateChange = vi.fn();
  render(
    <TegoSheet
      defaultDocument={source}
      mode="template"
      template={template}
      permissionStore={store}
      onTemplateChange={onTemplateChange}
    />,
  );

  fireEvent.click(await screen.findByRole('button', { name: 'Add value' }));
  expect(onTemplateChange).not.toHaveBeenCalled();

  act(() => {
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-template',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:view',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'sheet:view',
            target: { type: 'sheet', sheetId },
          },
          {
            action: 'template:bind',
            target: { type: 'sheet', sheetId },
          },
        ],
      }),
    );
  });
  fireEvent.click(screen.getByRole('button', { name: 'Add value' }));
  expect(onTemplateChange).toHaveBeenCalledTimes(1);
});

it('hides and restores the document live behind document:view and active sheet:view', async () => {
  const source = testDocument([{ name: 'Confidential' }]);
  const sheetId = source.workbook.sheets[0]!.id;
  const store = createPermissionStore();
  const rendered = render(<TegoSheet defaultDocument={source} permissionStore={store} />);

  expect(rendered.container.querySelector('canvas')).toBeNull();

  act(() => {
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-view',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:view',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'sheet:view',
            target: { type: 'sheet', sheetId },
          },
        ],
      }),
    );
  });
  await waitFor(() => expect(rendered.container.querySelector('canvas')).not.toBeNull());

  act(() => store.clear());
  await waitFor(() => expect(rendered.container.querySelector('canvas')).toBeNull());
});

it('requires print permission before entering preview rendering', async () => {
  const source = testDocument([{ name: 'Confidential' }]);
  const sheetId = source.workbook.sheets[0]!.id;
  const store = createPermissionStore();
  const template: SpreadsheetTemplate = {
    id: 'template-1' as never,
    name: 'Invoice',
    bindings: [],
    printProfiles: [],
  };
  store.replace(
    createPermissionSnapshot({
      revision: 'permission-view',
      actorId: 'actor-1',
      grants: [
        {
          action: 'document:view',
          target: { type: 'document', documentId: source.id },
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId },
        },
      ],
    }),
  );
  render(
    <TegoSheet document={source} mode="preview" template={template} permissionStore={store} />,
  );

  expect(screen.queryByLabelText('Template preview diagnostics')).toBeNull();

  act(() => {
    store.replace(
      createPermissionSnapshot({
        revision: 'permission-print',
        actorId: 'actor-1',
        grants: [
          {
            action: 'document:view',
            target: { type: 'document', documentId: source.id },
          },
          {
            action: 'sheet:view',
            target: { type: 'sheet', sheetId },
          },
          {
            action: 'print',
            target: { type: 'document', documentId: source.id },
          },
        ],
      }),
    );
  });
  expect((await screen.findByLabelText('Template preview diagnostics')).textContent).toMatch(
    /requires a template and deterministic render environment/u,
  );
});
