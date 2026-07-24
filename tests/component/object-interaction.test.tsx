import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument, TegoSheet, type TegoSheetHandle } from '../../src';
import type { SpreadsheetDocumentInput } from '../../src/document';
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

function objectDocument() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'object-interaction',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [],
          merges: [],
          rowCount: 10,
          columnCount: 5,
          objects: [
            {
              id: 'shape-1',
              kind: 'shape',
              anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
              zIndex: 1,
              locked: false,
              templateRepeat: 'shared',
              shape: 'rectangle',
              style: { fill: '#ffeecc' },
              accessibility: { name: 'Revenue chart', description: 'Quarterly revenue' },
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
  } as unknown as SpreadsheetDocumentInput);
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return parsed.document;
}

it('commits keyboard object movement through controller history and undoes it', async () => {
  const ref = createRef<TegoSheetHandle>();
  const rendered = render(<TegoSheet ref={ref} defaultDocument={objectDocument()} />);
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  fireEvent(window, new Event('resize'));

  const object = await rendered.findByRole('option', { name: 'Revenue chart' });
  fireEvent.keyDown(object, { key: 'ArrowRight' });
  await waitFor(() =>
    expect(ref.current!.getDocument().workbook.sheets[0]?.objects[0]?.anchor).toMatchObject({
      rect: { x: 11, y: 20, width: 40, height: 30 },
    }),
  );

  act(() => ref.current!.undo());
  expect(ref.current!.getDocument().workbook.sheets[0]?.objects[0]?.anchor).toMatchObject({
    rect: { x: 10, y: 20, width: 40, height: 30 },
  });
});
