import { act, cleanup, fireEvent, render, within } from '@testing-library/react';
import type { ComponentProps } from 'react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { TegoSheet, WorkbookData } from 'tego-sheet';
import { PREVIEW_EVENT_LIMIT } from '../../demo/src/workbench-model';

type SheetProps = ComponentProps<typeof TegoSheet>;

const sheetMock = vi.hoisted(() => ({
  currentProps: undefined as SheetProps | undefined,
  handleValue: [{ name: 'Exported' }] as WorkbookData,
  mounts: 0,
}));

vi.mock('tego-sheet', async () => {
  const React = await import('react');

  const MockTegoSheet = React.forwardRef<unknown, SheetProps>((props, ref) => {
    sheetMock.currentProps = props;

    React.useEffect(() => {
      sheetMock.mounts += 1;
    }, []);

    React.useImperativeHandle(ref, () => ({
      getValue: () => sheetMock.handleValue,
    }));

    const workbook = props.value ?? props.defaultValue ?? [];

    return (
      <div
        data-testid="tego-sheet"
        data-locale={props.locale?.id ?? 'en'}
        data-mode={props.value === undefined ? 'uncontrolled' : 'controlled'}
        data-read-only={String(props.readOnly ?? false)}
        data-workbook={JSON.stringify(workbook)}
      />
    );
  });
  MockTegoSheet.displayName = 'MockTegoSheet';

  return { TegoSheet: MockTegoSheet };
});

import { App } from '../../demo/src/app';

function currentSheetProps(): SheetProps {
  if (sheetMock.currentProps === undefined) throw new Error('The TegoSheet mock has not rendered.');
  return sheetMock.currentProps;
}

function workbookFromBoundary(boundary: HTMLElement): unknown {
  return JSON.parse(boundary.dataset.workbook ?? 'null') as unknown;
}

beforeEach(() => {
  sheetMock.currentProps = undefined;
  sheetMock.handleValue = [{ name: 'Exported' }];
  sheetMock.mounts = 0;
});

afterEach(() => {
  cleanup();
});

describe('demo workbench', () => {
  it('renders one uncontrolled sheet and remounts once when mode becomes controlled', () => {
    const rendered = render(<App />);
    const boundaries = rendered.getAllByTestId('tego-sheet');

    expect(boundaries).toHaveLength(1);
    expect(boundaries[0]?.getAttribute('data-mode')).toBe('uncontrolled');
    expect(currentSheetProps().defaultValue).toBeDefined();
    expect(currentSheetProps().value).toBeUndefined();
    expect(sheetMock.mounts).toBe(1);

    act(() => currentSheetProps().onChange?.(
      [{ name: 'Uncontrolled internal update' }],
      { id: 'change-0', kind: 'cell', source: 'keyboard', sheet: 'sheet-1' as never },
    ));
    expect(currentSheetProps().defaultValue).toEqual([{ name: 'Uncontrolled internal update' }]);
    expect(currentSheetProps().value).toBeUndefined();
    expect(rendered.getByText(/workbook: uncontrolled internal update/i)).toBeTruthy();
    expect(sheetMock.mounts).toBe(1);

    fireEvent.change(rendered.getByRole('combobox', { name: 'Mode' }), {
      target: { value: 'controlled' },
    });

    expect(rendered.getByTestId('tego-sheet').getAttribute('data-mode')).toBe('controlled');
    expect(currentSheetProps().value).toEqual([{ name: 'Uncontrolled internal update' }]);
    expect(currentSheetProps().defaultValue).toBeUndefined();
    expect(sheetMock.mounts).toBe(2);

    act(() => currentSheetProps().onChange?.(
      [{ name: 'Controlled update' }],
      { id: 'change-1', kind: 'cell', source: 'keyboard', sheet: 'sheet-1' as never },
    ));
    expect(workbookFromBoundary(rendered.getByTestId('tego-sheet'))).toEqual([{ name: 'Controlled update' }]);
  });

  it('passes read-only and locale selections through public props', () => {
    const rendered = render(<App />);

    fireEvent.click(rendered.getByRole('checkbox', { name: 'Read only' }));
    expect(currentSheetProps().readOnly).toBe(true);
    expect(rendered.getByTestId('tego-sheet').getAttribute('data-read-only')).toBe('true');

    fireEvent.change(rendered.getByRole('combobox', { name: 'Locale' }), {
      target: { value: 'zh-CN' },
    });
    expect(currentSheetProps().locale?.id).toBe('zh-CN');
    expect(rendered.getByTestId('tego-sheet').getAttribute('data-locale')).toBe('zh-CN');

    fireEvent.change(rendered.getByRole('combobox', { name: 'Locale' }), {
      target: { value: 'en' },
    });
    expect(currentSheetProps().locale).toBeUndefined();
  });

  it('reports invalid JSON without replacing or remounting the workbook', () => {
    const rendered = render(<App />);
    const initialWorkbook = workbookFromBoundary(rendered.getByTestId('tego-sheet'));
    const initialMounts = sheetMock.mounts;

    fireEvent.click(rendered.getByRole('button', { name: 'Show JSON' }));
    const json = rendered.getByRole('textbox', { name: 'Workbook JSON' });
    fireEvent.change(json, { target: { value: '{ invalid' } });
    fireEvent.click(rendered.getByRole('button', { name: 'Import JSON' }));

    expect(rendered.getByRole('alert').textContent).toMatch(/invalid|json|unexpected/i);
    expect(workbookFromBoundary(rendered.getByTestId('tego-sheet'))).toEqual(initialWorkbook);
    expect(sheetMock.mounts).toBe(initialMounts);
  });

  it('imports, resets, and exports workbook JSON through the public handle', () => {
    const rendered = render(<App />);
    fireEvent.click(rendered.getByRole('button', { name: 'Show JSON' }));
    const json = rendered.getByRole('textbox', { name: 'Workbook JSON' });

    fireEvent.change(json, { target: { value: '[{"name":"Imported"}]' } });
    fireEvent.click(rendered.getByRole('button', { name: 'Import JSON' }));
    expect(workbookFromBoundary(rendered.getByTestId('tego-sheet'))).toEqual([{ name: 'Imported' }]);
    expect(sheetMock.mounts).toBe(2);

    fireEvent.click(rendered.getByRole('button', { name: 'Reset workbook' }));
    expect(workbookFromBoundary(rendered.getByTestId('tego-sheet'))).toEqual(
      expect.arrayContaining([expect.objectContaining({ name: 'Budget' })]),
    );
    expect(sheetMock.mounts).toBe(3);

    sheetMock.handleValue = [{ name: 'From handle', rows: { len: 1 } }];
    fireEvent.click(rendered.getByRole('button', { name: 'Export JSON' }));
    expect((json as HTMLTextAreaElement).value).toBe('[\n  {\n    "name": "From handle",\n    "rows": {\n      "len": 1\n    }\n  }\n]');
  });

  it('records callback events newest-first and bounds the event log', () => {
    const rendered = render(<App />);

    act(() => {
      for (let index = 0; index < PREVIEW_EVENT_LIMIT; index += 1) {
        currentSheetProps().onChange?.(
          [{ name: `Change ${index}` }],
          { id: `change-${index}`, kind: 'cell', source: 'keyboard', sheet: 'sheet-1' as never },
        );
      }
      currentSheetProps().onActiveSheetChange?.({ sheet: 'sheet-2' as never, index: 1, source: 'sheet-tabs' });
      currentSheetProps().onSelectionChange?.({
        sheet: 'sheet-2' as never,
        range: { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
        active: { row: 1, column: 1 },
      });
      currentSheetProps().onError?.({ code: 'RENDER_FAILED', message: 'Canvas unavailable', recoverable: true });
    });

    fireEvent.click(rendered.getByRole('button', { name: 'Show events' }));
    const log = rendered.getByRole('log');
    const items = within(log).getAllByRole('listitem');

    expect(items).toHaveLength(PREVIEW_EVENT_LIMIT);
    expect(items[0]?.textContent).toMatch(/spreadsheet error.*canvas unavailable/i);
    expect(log.textContent).toMatch(/selection changed/i);
    expect(log.textContent).toMatch(/active sheet changed/i);
    expect(log.textContent).toMatch(/workbook changed/i);
    expect(log.textContent).not.toContain('"id":"change-0"');
  });

  it('provides disclosure state and keeps primary status available when controls collapse', () => {
    const rendered = render(<App />);
    const jsonDisclosure = rendered.getByRole('button', { name: 'Show JSON' });
    const eventDisclosure = rendered.getByRole('button', { name: 'Show events' });
    const collapse = rendered.getByRole('button', { name: 'Collapse controls' });

    expect(jsonDisclosure.getAttribute('aria-expanded')).toBe('false');
    expect(eventDisclosure.getAttribute('aria-expanded')).toBe('false');
    expect(collapse.getAttribute('aria-expanded')).toBe('true');

    fireEvent.click(jsonDisclosure);
    expect(rendered.getByRole('button', { name: 'Hide JSON' }).getAttribute('aria-expanded')).toBe('true');
    fireEvent.click(rendered.getByRole('button', { name: 'Hide JSON' }));
    fireEvent.click(eventDisclosure);
    expect(rendered.getByRole('button', { name: 'Hide events' }).getAttribute('aria-expanded')).toBe('true');
    fireEvent.click(rendered.getByRole('button', { name: 'Hide events' }));
    fireEvent.click(collapse);

    expect(rendered.getByRole('heading', { name: /tego-sheet workbench/i })).toBeTruthy();
    expect(rendered.getByText(/mode: uncontrolled/i)).toBeTruthy();
    expect(rendered.getByText(/workbook: budget/i)).toBeTruthy();
    expect(rendered.getByRole('button', { name: 'Expand controls' }).getAttribute('aria-expanded')).toBe('false');
    expect(rendered.queryByRole('combobox', { name: 'Mode' })).toBeNull();
    expect(rendered.queryByRole('button', { name: 'Reset workbook' })).toBeNull();
  });

  it('exposes stable fullscreen layout hooks and disclosure relationships', () => {
    const rendered = render(<App />);
    const shell = rendered.container.querySelector('.preview-shell');
    const controls = rendered.container.querySelector('.preview-controls');
    const workspace = rendered.container.querySelector('.preview-workspace');
    const jsonDisclosure = rendered.getByRole('button', { name: 'Show JSON' });
    const eventDisclosure = rendered.getByRole('button', { name: 'Show events' });
    const collapseDisclosure = rendered.getByRole('button', { name: 'Collapse controls' });

    expect(shell).toBeTruthy();
    expect(controls).toBeTruthy();
    expect(workspace).toBeTruthy();

    fireEvent.click(jsonDisclosure);
    const jsonPanel = rendered.container.querySelector('.preview-json-panel');
    expect(jsonPanel).toBeTruthy();
    expect(jsonDisclosure.getAttribute('aria-expanded')).toBe('true');
    expect(jsonDisclosure.getAttribute('aria-controls')).toBe(jsonPanel?.id);

    fireEvent.click(eventDisclosure);
    const eventsPanel = rendered.container.querySelector('.preview-events-panel');
    expect(eventsPanel).toBeTruthy();
    expect(eventDisclosure.getAttribute('aria-expanded')).toBe('true');
    expect(eventDisclosure.getAttribute('aria-controls')).toBe(eventsPanel?.id);

    const secondaryControlsId = collapseDisclosure.getAttribute('aria-controls');
    expect(collapseDisclosure.getAttribute('aria-expanded')).toBe('true');
    expect(secondaryControlsId).toBeTruthy();
    expect(rendered.container.querySelector(`#${secondaryControlsId}`)).toBeTruthy();

    fireEvent.click(collapseDisclosure);
    expect(rendered.getByRole('button', { name: 'Expand controls' }).getAttribute('aria-expanded')).toBe('false');
  });
});
