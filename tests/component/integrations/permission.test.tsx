import { act, cleanup, render, waitFor } from '@testing-library/react';
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
