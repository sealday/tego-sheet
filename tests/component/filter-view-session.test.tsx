import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument, TegoSheet, type TegoSheetHandle } from '../../src';
import { createCanvasHarness } from '../helpers/canvas-harness';

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

it('activates a saved view as session-only screen and accessibility state', async () => {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'view-component',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [
            { row: 0, column: 0, cell: { input: { type: 'string', value: 'status' } } },
            { row: 1, column: 0, cell: { input: { type: 'string', value: 'keep' } } },
            { row: 2, column: 0, cell: { input: { type: 'string', value: 'drop' } } },
          ],
          merges: [],
          filterViews: [
            {
              id: 'kept',
              name: 'Kept',
              range: {
                sheetId: 'sheet-1',
                start: { row: 0, column: 0 },
                end: { row: 2, column: 0 },
              },
              sorts: [],
              filters: [{ column: 0, operator: 'equal', value: 'keep' }],
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
  const ref = createRef<TegoSheetHandle>();
  const onDocumentChange = vi.fn();
  const rendered = render(
    <TegoSheet ref={ref} defaultDocument={parsed.document} onDocumentChange={onDocumentChange} />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  fireEvent(window, new Event('resize'));
  await waitFor(() => expect(rendered.getByText('drop')).toBeTruthy());

  ref.current!.activateFilterView(
    ref.current!.getDocument().workbook.sheets[0]!.id as never,
    'kept',
  );
  await waitFor(() => expect(rendered.queryByText('drop')).toBeNull());
  expect(onDocumentChange).not.toHaveBeenCalled();
  expect(ref.current!.getDocument().workbook.sheets[0]).not.toHaveProperty('activeViewId');

  ref.current!.deactivateFilterView(ref.current!.getDocument().workbook.sheets[0]!.id as never);
  await waitFor(() => expect(rendered.getByText('drop')).toBeTruthy());
});
