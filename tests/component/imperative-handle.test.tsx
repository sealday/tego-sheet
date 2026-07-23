import { act, cleanup, render, waitFor } from '@testing-library/react';
import { createRef, startTransition, StrictMode, Suspense } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet, type TegoSheetHandle } from '../../src';
import { createCanvasHarness } from '../helpers/canvas-harness';
import { legacyProjection, testDocument } from '../helpers/workbook-builders';

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

it('exposes every approved command and isolated query through one stable handle', async () => {
  const ref = createRef<TegoSheetHandle>();
  const onActiveSheetChange = vi.fn();
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={testDocument([{ name: 'A', rows: { len: 2 }, cols: { len: 2 } }])}
      options={{ defaultStyle: { color: '#123456' } }}
      onActiveSheetChange={onActiveSheetChange}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const handle = ref.current!;
  const firstId = onActiveSheetChange.mock.lastCall?.[0].sheet ?? handle.addSheet('temporary');
  if (legacyProjection(handle.getDocument()).length > 1) handle.deleteSheet(firstId);
  const sheet = handle.addSheet('B');

  act(() => {
    handle.setCellText({ sheet, row: 0, column: 0 }, 'value');
    handle.renameSheet(sheet, 'Renamed');
    handle.activateSheet(sheet);
  });
  expect(handle.getCell({ sheet, row: 0, column: 0 })?.text).toBe('value');
  expect(handle.getCellStyle({ sheet, row: 0, column: 0 })).toMatchObject({ color: '#123456' });
  const value = legacyProjection(handle.getDocument()) as Array<{ name?: string }>;
  value[0]!.name = 'caller mutation';
  expect(legacyProjection(handle.getDocument())[0]?.name).not.toBe('caller mutation');
  expect(handle.validate()).toEqual({ valid: true, issues: [] });

  act(() => handle.undo());
  expect(legacyProjection(handle.getDocument()).at(-1)?.name).toBe('B');
  act(() => handle.redo());
  expect(legacyProjection(handle.getDocument()).at(-1)?.name).toBe('Renamed');
  handle.focus();
  expect(document.activeElement).toBe(rendered.container.querySelector('[data-tego-sheet]'));
  expect(() => handle.recalculateLayout()).not.toThrow();

  rendered.rerender(
    <TegoSheet
      ref={ref}
      defaultDocument={testDocument([])}
      options={{ defaultStyle: { color: 'red' } }}
    />,
  );
  expect(ref.current).toBe(handle);
});

it('invalidates stale sheet IDs and clips active index silently on external replacement', async () => {
  const ref = createRef<TegoSheetHandle>();
  const onChange = vi.fn();
  const onActiveSheetChange = vi.fn();
  const onSelectionChange = vi.fn();
  const first = [{ name: 'A' }, { name: 'B' }];
  const rendered = render(
    <TegoSheet
      ref={ref}
      document={testDocument(first)}
      initialActiveSheetIndex={1}
      onDocumentChange={onChange}
      onActiveSheetChange={onActiveSheetChange}
      onSelectionChange={onSelectionChange}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const stale = ref.current!.addSheet('optimistic');
  onChange.mockClear();

  rendered.rerender(
    <TegoSheet
      ref={ref}
      document={testDocument([{ name: 'Replacement' }])}
      initialActiveSheetIndex={99}
      onDocumentChange={onChange}
      onActiveSheetChange={onActiveSheetChange}
      onSelectionChange={onSelectionChange}
    />,
  );
  await waitFor(() =>
    expect(legacyProjection(ref.current!.getDocument())[0]?.name).toBe('Replacement'),
  );
  expect(() => ref.current!.getCell({ sheet: stale, row: 0, column: 0 })).toThrowError(
    expect.objectContaining({ code: 'INVALID_COMMAND' }),
  );
  expect(onChange).not.toHaveBeenCalled();
  expect(onActiveSheetChange).not.toHaveBeenCalled();
  expect(onSelectionChange).not.toHaveBeenCalled();
});

it.each([
  ['empty', []],
  ['non-empty', [{ name: 'A' }]],
] as const)('focuses the %s sheet when its callback ref first receives the handle', (_, value) => {
  let handle: TegoSheetHandle | null = null;
  let focused: Element | null = null;
  const rendered = render(
    <TegoSheet
      defaultDocument={testDocument(value)}
      ref={(next) => {
        if (next === null) return;
        handle = next;
        next.focus();
        focused = document.activeElement;
      }}
    />,
  );
  const root = rendered.container.querySelector('[data-tego-sheet]');

  expect(handle).not.toBeNull();
  expect(focused).toBe(root);
  expect(document.activeElement).toBe(root);
});

it('keeps callback-ref roots isolated across StrictMode teardown and unmount', () => {
  let handle: TegoSheetHandle | null = null;
  const attachments: Array<TegoSheetHandle | null> = [];
  const rendered = render(
    <StrictMode>
      <TegoSheet
        defaultDocument={testDocument([])}
        ref={(next) => {
          attachments.push(next);
          if (next === null) return;
          handle = next;
          next.focus();
        }}
      />
    </StrictMode>,
  );
  const root = rendered.container.querySelector('[data-tego-sheet]');
  expect(document.activeElement).toBe(root);

  rendered.unmount();

  expect(attachments.at(-1)).toBeNull();
  expect(() => handle!.focus()).toThrowError(expect.objectContaining({ code: 'INVALID_COMMAND' }));
  expect(document.activeElement).not.toBe(root);
});

it('does not attach an imperative root from an aborted render', () => {
  const suspended = new Promise<never>(() => undefined);
  let committedHandle: TegoSheetHandle | null = null;
  let pendingAttachments = 0;

  function SuspendAfterSheet(props: { readonly active: boolean }) {
    if (props.active) throw suspended;
    return null;
  }

  function Host(props: { readonly pending: boolean }) {
    return (
      <Suspense fallback={<output data-suspended="" />}>
        <TegoSheet
          key={props.pending ? 'pending' : 'committed'}
          defaultDocument={testDocument([{ name: props.pending ? 'Pending' : 'Committed' }])}
          ref={(next) => {
            if (next === null) return;
            if (props.pending) pendingAttachments += 1;
            else committedHandle = next;
          }}
        />
        <SuspendAfterSheet active={props.pending} />
      </Suspense>
    );
  }

  const rendered = render(<Host pending={false} />);
  const committedRoot = rendered.container.querySelector('[data-tego-sheet]');
  act(() => {
    startTransition(() => rendered.rerender(<Host pending />));
  });
  committedHandle!.focus();

  expect(pendingAttachments).toBe(0);
  expect(rendered.container.querySelector('[data-tego-sheet]')).toBe(committedRoot);
  expect(document.activeElement).toBe(committedRoot);
});
