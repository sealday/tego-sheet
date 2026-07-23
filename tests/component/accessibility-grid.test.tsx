import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { afterEach, expect, it, vi } from 'vitest';
import { AccessibilityGrid } from '../../src/react/accessibility/accessibility-grid';
import type { CellPresentation } from '../../src/presentation';
import { TegoSheet } from '../../src';
import { createCanvasHarness } from '../helpers/canvas-harness';
import { testDocument } from '../helpers/workbook-builders';

afterEach(cleanup);

function presentation({ row, column }: { row: number; column: number }): CellPresentation {
  const text = row === 2 && column === 3 ? '#DIV/0!' : `R${row + 1}C${column + 1}`;
  const invalid = row === 2 && column === 3;
  return {
    address: { sheetId: 'sheet-1' as never, row, column },
    value: invalid ? { type: 'error', value: '#DIV/0!' } : { type: 'string', value: text },
    formattedText: text,
    style: {
      color: '#000000',
      backgroundColor: '#ffffff',
      fontFamily: 'Arial',
      fontSize: 10,
      bold: false,
      italic: false,
      horizontalAlign: 'left',
      verticalAlign: 'middle',
      wrap: false,
    },
    validation: invalid ? { status: 'error', message: 'Division by zero' } : { status: 'valid' },
    annotations: [],
    visibility: { hidden: false, printable: true },
    accessibility: {
      label: invalid ? '#DIV/0!, Division by zero' : text,
      readOnly: false,
      invalid,
      ...(invalid ? { description: 'Division by zero' } : {}),
    },
  };
}

it('renders only viewport plus bounded overscan for very large sheets', () => {
  const rendered = render(
    <AccessibilityGrid
      rowCount={1_000_000}
      columnCount={16_384}
      viewport={{ rowStart: 100, rowEnd: 119, columnStart: 40, columnEnd: 49 }}
      overscan={2}
      activeCell={{ row: 105, column: 42 }}
      selection={{ start: { row: 100, column: 40 }, end: { row: 119, column: 49 } }}
      resolvePresentation={presentation}
    />,
  );

  const cells = rendered.container.querySelectorAll('[role="gridcell"]');
  expect(cells.length).toBeLessThan(500);
  expect(cells.length).toBe(24 * 14);
  expect(rendered.getByRole('grid').getAttribute('aria-rowcount')).toBe('1000000');
  expect(rendered.getByRole('grid').getAttribute('aria-colcount')).toBe('16384');
});

it('keeps one active gridcell, exposes error semantics and summarizes large selections', () => {
  const rendered = render(
    <AccessibilityGrid
      rowCount={100}
      columnCount={100}
      viewport={{ rowStart: 0, rowEnd: 9, columnStart: 0, columnEnd: 9 }}
      activeCell={{ row: 2, column: 3 }}
      selection={{ start: { row: 0, column: 0 }, end: { row: 99, column: 99 } }}
      resolvePresentation={presentation}
    />,
  );

  const active = rendered.getByRole('gridcell', { name: '#DIV/0!, Division by zero' });
  expect(active.getAttribute('tabindex')).toBe('0');
  expect(active.getAttribute('aria-invalid')).toBe('true');
  expect(
    [...rendered.container.querySelectorAll('[role="gridcell"]')].filter(
      (cell) => cell.getAttribute('tabindex') === '0',
    ),
  ).toHaveLength(1);
  expect(rendered.getByText('Selected 10000 cells').getAttribute('aria-live')).toBe('polite');
});

it('activates and edits semantic cells, then restores focus when the editor closes', () => {
  const onActivate = vi.fn();
  const onRequestEdit = vi.fn();
  const rendered = render(
    <AccessibilityGrid
      rowCount={20}
      columnCount={20}
      viewport={{ rowStart: 0, rowEnd: 9, columnStart: 0, columnEnd: 9 }}
      activeCell={{ row: 1, column: 1 }}
      selection={{ start: { row: 1, column: 1 }, end: { row: 1, column: 1 } }}
      resolvePresentation={presentation}
      editorOpen
      onActivate={onActivate}
      onRequestEdit={onRequestEdit}
    />,
  );
  const target = rendered.getByRole('gridcell', { name: 'R2C2' });

  fireEvent.click(target);
  fireEvent.doubleClick(target);
  expect(onActivate).toHaveBeenCalledWith({ row: 1, column: 1 });
  expect(onRequestEdit).toHaveBeenCalledWith({ row: 1, column: 1 });

  rendered.rerender(
    <AccessibilityGrid
      rowCount={20}
      columnCount={20}
      viewport={{ rowStart: 0, rowEnd: 9, columnStart: 0, columnEnd: 9 }}
      activeCell={{ row: 1, column: 1 }}
      selection={{ start: { row: 1, column: 1 }, end: { row: 1, column: 1 } }}
      resolvePresentation={presentation}
      editorOpen={false}
      onActivate={onActivate}
      onRequestEdit={onRequestEdit}
    />,
  );

  expect(document.activeElement).toBe(rendered.getByRole('gridcell', { name: 'R2C2' }));
});

it('mounts the bounded semantic grid beside the Canvas surface', async () => {
  const context = createCanvasHarness().canvas.getContext('2d');
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => context);
  vi.stubGlobal(
    'requestAnimationFrame',
    vi.fn((callback: FrameRequestCallback) => {
      callback(0);
      return 1;
    }),
  );
  vi.stubGlobal('cancelAnimationFrame', vi.fn());

  const rendered = render(
    <TegoSheet
      defaultDocument={testDocument([
        {
          rows: {
            len: 100_000,
            0: { cells: { 0: { text: 'shared text' } } },
          },
          cols: { len: 10_000 },
        },
      ])}
    />,
  );

  await waitFor(() => expect(rendered.queryByRole('grid')).not.toBeNull());
  expect(rendered.container.querySelector('canvas')).not.toBeNull();
  expect(rendered.getByRole('gridcell', { name: 'shared text' }).textContent).toBe('shared text');
  expect(rendered.container.querySelectorAll('[role="gridcell"]').length).toBeLessThan(500);
});
