import { cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import {
  forwardRef,
  useImperativeHandle,
  useState,
  type ForwardedRef,
  type ReactNode,
} from 'react';
import type { TegoSheetHandle, WorkbookData } from 'tego-sheet';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';

interface SheetDoubleProps {
  readonly value?: unknown;
  readonly defaultValue?: unknown;
  readonly locale?: { readonly id: string };
  readonly toolbar?: ((props: Record<string, unknown>) => ReactNode) | string | false;
  readonly sheetTabs?: ((props: Record<string, unknown>) => ReactNode) | string | false;
  readonly onChange?: (value: unknown, change: unknown) => void;
  readonly onActiveSheetChange?: (event: unknown) => void;
  readonly onSelectionChange?: (event: unknown) => void;
  readonly onCellEdit?: (event: unknown) => void;
  readonly onPaste?: (event: unknown) => void;
  readonly onError?: (event: unknown) => void;
}

const sheetMock = vi.hoisted(() => ({
  failRender: false,
  mountSequence: 0,
  toolbarActions: [] as unknown[],
  tabActions: [] as string[],
  recordProps: vi.fn(),
}));

vi.mock('tego-sheet', async () => {
  const React = await import('react');

  const TegoSheet = forwardRef(function TegoSheetMock(
    props: SheetDoubleProps,
    ref: ForwardedRef<TegoSheetHandle>,
  ) {
    if (sheetMock.failRender) {
      throw new Error('unexpected render failure');
    }

    const [mountId] = useState(() => String(++sheetMock.mountSequence));
    sheetMock.recordProps(props);
    const workbook = (props.value ?? props.defaultValue ?? []) as WorkbookData;
    useImperativeHandle(ref, () => ({
      focus: () => undefined,
      getValue: () => workbook,
      getCell: () => null,
      getCellStyle: () => ({}),
      setCellText: () => undefined,
      addSheet: () => 'mock-sheet' as ReturnType<TegoSheetHandle['addSheet']>,
      deleteSheet: () => undefined,
      renameSheet: () => undefined,
      activateSheet: () => undefined,
      undo: () => undefined,
      redo: () => undefined,
      validate: () => ({ valid: true, issues: [] }),
      print: () => undefined,
      recalculateLayout: () => undefined,
    }));

    const toolbar =
      typeof props.toolbar === 'function'
        ? props.toolbar({
            selection: null,
            activeStyle: {},
            readOnly: false,
            canUndo: true,
            canRedo: false,
            merged: false,
            frozen: false,
            disabledActions: new Set(),
            execute: (action: unknown) => sheetMock.toolbarActions.push(action),
          })
        : null;
    const sheetTabs =
      typeof props.sheetTabs === 'function'
        ? props.sheetTabs({
            sheets: [{ id: 'mock-sheet', index: 0, name: 'Roadmap' }],
            activeSheet: 'mock-sheet',
            readOnly: false,
            add: () => sheetMock.tabActions.push('add'),
            delete: () => sheetMock.tabActions.push('delete'),
            rename: () => sheetMock.tabActions.push('rename'),
            activate: () => sheetMock.tabActions.push('activate'),
          })
        : null;

    return React.createElement(
      'div',
      {
        'data-testid': 'tego-sheet-double',
        'data-locale': props.locale?.id ?? 'default',
        'data-mount-id': mountId,
        'data-ownership': props.value === undefined ? 'uncontrolled' : 'controlled',
      },
      toolbar,
      sheetTabs,
      React.createElement(
        'button',
        {
          type: 'button',
          onClick: () =>
            props.onChange?.(
              [{ name: 'Accepted callback value', rows: { len: sheetMock.mountSequence } }],
              { id: `change-${sheetMock.mountSequence}`, kind: 'cell', source: 'keyboard' },
            ),
        },
        'Commit mock change',
      ),
      React.createElement(
        'button',
        {
          type: 'button',
          onClick: () =>
            props.onError?.({
              code: 'CLIPBOARD_DENIED',
              message: 'Clipboard permission denied',
              recoverable: true,
            }),
        },
        'Emit mock error',
      ),
    );
  });

  return { TegoSheet };
});

vi.mock('tego-sheet/locales/en', () => ({ en: { id: 'en', messages: {} } }));
vi.mock('tego-sheet/locales/zh-cn', () => ({ zhCN: { id: 'zh-CN', messages: {} } }));
vi.mock('tego-sheet/locales/de', () => ({ de: { id: 'de', messages: {} } }));
vi.mock('tego-sheet/locales/nl', () => ({ nl: { id: 'nl', messages: {} } }));

beforeEach(() => {
  sheetMock.failRender = false;
  sheetMock.mountSequence = 0;
  sheetMock.toolbarActions.length = 0;
  sheetMock.tabActions.length = 0;
  sheetMock.recordProps.mockClear();
  window.history.replaceState({}, '', '/tego-sheet/playground');
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

async function renderPlayground() {
  const { Playground } = await import('../../website/src/components/playground/playground');
  return render(<Playground />);
}

function latestSheetProps(): SheetDoubleProps | undefined {
  return sheetMock.recordProps.mock.lastCall?.[0] as SheetDoubleProps | undefined;
}

it('selects Controlled from the initial mode query', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?mode=controlled');

  await renderPlayground();

  expect((screen.getByRole('radio', { name: 'Controlled' }) as HTMLInputElement).checked).toBe(
    true,
  );
  expect(screen.getByTestId('tego-sheet-double').getAttribute('data-ownership')).toBe('controlled');
  expect(latestSheetProps()?.value).toBeDefined();
  expect(latestSheetProps()?.defaultValue).toBeUndefined();
});

it('canonicalizes an invalid mode while preserving the path and unrelated parameters', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?theme=dark&mode=nope');
  const replaceState = vi.spyOn(window.history, 'replaceState');

  await renderPlayground();

  expect((screen.getByRole('radio', { name: 'Uncontrolled' }) as HTMLInputElement).checked).toBe(
    true,
  );
  expect(replaceState).toHaveBeenCalledWith(
    window.history.state,
    '',
    '/tego-sheet/playground?theme=dark&mode=uncontrolled',
  );
});

it('pushes a selected mode, changes the keyed preset boundary, and clears old events', async () => {
  await renderPlayground();
  fireEvent.click(screen.getByRole('button', { name: 'Commit mock change' }));
  expect(screen.getAllByRole('listitem', { name: /event/i })).toHaveLength(1);
  const oldKey = screen.getByTestId('preset-boundary').getAttribute('data-preset-key');
  const pushState = vi.spyOn(window.history, 'pushState');

  fireEvent.click(screen.getByRole('radio', { name: 'Controlled' }));

  expect(pushState).toHaveBeenCalledWith(
    window.history.state,
    '',
    '/tego-sheet/playground?mode=controlled',
  );
  expect(screen.getByTestId('preset-boundary').getAttribute('data-preset-key')).not.toBe(oldKey);
  expect(screen.queryAllByRole('listitem', { name: /event/i })).toHaveLength(0);
});

it('restores the matching mode when browser history emits popstate', async () => {
  await renderPlayground();
  window.history.pushState({}, '', '/tego-sheet/playground?mode=locales');

  fireEvent(window, new PopStateEvent('popstate'));

  expect((screen.getByRole('radio', { name: 'Locales' }) as HTMLInputElement).checked).toBe(true);
  expect(screen.getByLabelText('Locale')).toBeTruthy();
});

it('Reset mode recreates the current fixture without reloading the page', async () => {
  await renderPlayground();
  const firstMount = screen.getByTestId('tego-sheet-double').getAttribute('data-mount-id');

  fireEvent.click(screen.getByRole('button', { name: 'Reset mode' }));

  expect(screen.getByTestId('tego-sheet-double').getAttribute('data-mount-id')).not.toBe(
    firstMount,
  );
  expect(screen.getByRole('status').textContent).toContain('Uncontrolled reset');
});

it('accepts controlled callbacks into the displayed workbook JSON', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?mode=controlled');
  await renderPlayground();

  fireEvent.click(screen.getByRole('button', { name: 'Commit mock change' }));

  expect(screen.getByLabelText('Workbook JSON').textContent).toContain('Accepted callback value');
  expect(screen.getByTestId('tego-sheet-double').getAttribute('data-ownership')).toBe('controlled');
});

it('retains only the newest 50 callback events in newest-first order', async () => {
  await renderPlayground();
  const commit = screen.getByRole('button', { name: 'Commit mock change' });

  for (let index = 0; index < 55; index += 1) fireEvent.click(commit);

  const events = screen.getAllByRole('listitem', { name: /event/i });
  expect(events).toHaveLength(50);
  expect(events[0]?.textContent).toContain('#55 onChange');
  expect(events.at(-1)?.textContent).toContain('#6 onChange');
});

it('switches among all four public locale dictionaries per instance', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?mode=locales');
  await renderPlayground();
  const locale = screen.getByLabelText('Locale');

  expect(screen.getByTestId('tego-sheet-double').getAttribute('data-locale')).toBe('en');
  for (const [label, id] of [
    ['简体中文', 'zh-CN'],
    ['Deutsch', 'de'],
    ['Nederlands', 'nl'],
    ['English', 'en'],
  ]) {
    fireEvent.change(locale, { target: { value: id } });
    expect(screen.getByTestId('tego-sheet-double').getAttribute('data-locale')).toBe(id);
    expect(latestSheetProps()?.locale?.id).toBe(id);
    expect(within(locale).getByRole('option', { name: label })).toBeTruthy();
  }
});

it('renders custom chrome from typed public slot props and dispatches public actions', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?mode=custom-chrome');
  await renderPlayground();

  fireEvent.click(screen.getByRole('button', { name: 'Bold selection' }));
  fireEvent.click(screen.getByRole('button', { name: 'Add demo sheet' }));

  expect(sheetMock.toolbarActions).toEqual([
    { type: 'set-style', patch: { font: { bold: true } } },
  ]);
  expect(sheetMock.tabActions).toEqual(['add']);
  expect(typeof latestSheetProps()?.toolbar).toBe('function');
  expect(typeof latestSheetProps()?.sheetTabs).toBe('function');
});

it('records structured onError payloads as public events', async () => {
  await renderPlayground();

  fireEvent.click(screen.getByRole('button', { name: 'Emit mock error' }));

  const event = screen.getByRole('listitem', { name: /event 1/i });
  expect(event.textContent).toContain('onError');
  expect(event.textContent).toContain('CLIPBOARD_DENIED');
});

it('offers Reset and Reload after an unexpected render failure without exposing a stack', async () => {
  await renderPlayground();
  const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
  sheetMock.failRender = true;
  fireEvent.click(screen.getByRole('radio', { name: 'Controlled' }));

  expect(screen.getByRole('alert').textContent).toContain(
    'The playground could not render this preset',
  );
  expect(screen.queryByText(/unexpected render failure/)).toBeNull();
  expect(screen.getByRole('button', { name: 'Reset' })).toBeTruthy();
  expect(screen.getByRole('button', { name: 'Reload' })).toBeTruthy();

  sheetMock.failRender = false;
  fireEvent.click(screen.getByRole('button', { name: 'Reset' }));
  expect((screen.getByRole('radio', { name: 'Uncontrolled' }) as HTMLInputElement).checked).toBe(
    true,
  );
  expect(screen.getByTestId('tego-sheet-double')).toBeTruthy();
  expect(consoleError).toHaveBeenCalled();
});

it('copies formatted JSON and announces the result', async () => {
  const writeText = vi.fn().mockResolvedValue(undefined);
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: { writeText },
  });
  await renderPlayground();

  fireEvent.click(screen.getByRole('button', { name: 'Copy JSON' }));

  expect(writeText).toHaveBeenCalledWith(expect.stringContaining('"name": "Budget"'));
  await waitFor(() =>
    expect(screen.getByRole('status').textContent).toContain('Workbook JSON copied'),
  );
});
