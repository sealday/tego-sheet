import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { createRef, useState } from 'react';
import { flushSync } from 'react-dom';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import {
  parseSpreadsheetDocument,
  TegoSheet,
  type SpreadsheetDocument,
  type TegoSheetHandle,
} from '../../src';
import type {
  ValidationEngine,
  ValidationResult as AdvancedValidationResult,
} from '../../src/validation';
import type { WorkbookInput } from '../../src/core';
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

function sizeRoot(root: HTMLElement, width = 500, height = 300): void {
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: width },
    clientHeight: { configurable: true, value: height },
  });
  fireEvent(window, new Event('resize'));
}

function validatedDocument(behavior: 'reject' | 'warn' = 'reject'): SpreadsheetDocument {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'validated-editor',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'number', value: 1 }, validationId: 'amount' },
            },
          ],
          merges: [],
        },
      ],
      styles: [],
      validations: [
        {
          id: 'amount',
          value: {
            id: 'amount',
            type: 'number',
            predicate: { operator: 'between', minimum: 0, maximum: 100 },
            behavior,
            allowBlank: false,
          },
        },
      ],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return parsed.document;
}

function deferredValidation(): {
  readonly engine: ValidationEngine;
  readonly release: (result: AdvancedValidationResult) => void;
} {
  let release!: (result: AdvancedValidationResult) => void;
  const result = new Promise<AdvancedValidationResult>((resolve) => {
    release = resolve;
  });
  return { engine: { validate: () => result }, release };
}

it('validates editor commits asynchronously before one controlled callback sequence', async () => {
  const ref = createRef<TegoSheetHandle>();
  const order: string[] = [];
  const validation = deferredValidation();
  function Controlled() {
    const [document, setDocument] = useState(validatedDocument);
    return (
      <TegoSheet
        ref={ref}
        document={document}
        validationEngine={validation.engine}
        onDocumentChange={(next) => {
          order.push('change');
          setDocument(next);
        }}
        onCellEdit={() => order.push('cell-edit')}
        onSelectionChange={() => order.push('selection')}
      />
    );
  }
  const rendered = render(<Controlled />);
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: '12' } });
  fireEvent.keyDown(editor, { key: 'Enter' });
  expect(rendered.getByRole('textbox', { name: /cell editor/i })).toBe(editor);
  expect(order).toEqual([]);

  validation.release({ status: 'accepted', diagnostics: [] });
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());
  expect(order).toEqual(['change', 'cell-edit', 'selection']);
  expect(legacyProjection(ref.current!.getDocument())[0]?.rows?.['0']).toMatchObject({
    cells: { 0: { text: '12' } },
  });
});

it('gates imperative cell edits through the same asynchronous validation path', async () => {
  const ref = createRef<TegoSheetHandle>();
  const validation = deferredValidation();
  render(
    <TegoSheet
      ref={ref}
      defaultDocument={validatedDocument()}
      validationEngine={validation.engine}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());

  ref.current!.setCellText(
    { sheet: ref.current!.getDocument().workbook.sheets[0]!.id as never, row: 0, column: 0 },
    '12',
  );
  expect(legacyProjection(ref.current!.getDocument())[0]?.rows?.['0']).toMatchObject({
    cells: { 0: { type: 'number', value: 1 } },
  });

  validation.release({ status: 'accepted', diagnostics: [] });
  await waitFor(() =>
    expect(legacyProjection(ref.current!.getDocument())[0]?.rows?.['0']).toMatchObject({
      cells: { 0: { text: '12' } },
    }),
  );
});

it('aborts pending imperative validation when the component is disposed', async () => {
  const ref = createRef<TegoSheetHandle>();
  let observedSignal: AbortSignal | undefined;
  const validation = deferredValidation();
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={validatedDocument()}
      validationEngine={{
        validate: (request) => {
          observedSignal = request.signal;
          return validation.engine.validate(request);
        },
      }}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());

  ref.current!.setCellText(
    { sheet: ref.current!.getDocument().workbook.sheets[0]!.id as never, row: 0, column: 0 },
    '12',
  );
  expect(observedSignal?.aborted).toBe(false);
  rendered.unmount();
  expect(observedSignal?.aborted).toBe(true);
  validation.release({ status: 'accepted', diagnostics: [] });
  await Promise.resolve();
});

it('requires explicit host confirmation before committing a warning-mode editor value', async () => {
  const ref = createRef<TegoSheetHandle>();
  const confirmValidationWarning = vi.fn(() => true);
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={validatedDocument('warn')}
      validationEngine={{
        validate: async () => ({
          status: 'warning',
          code: 'VALIDATION_REJECTED',
          diagnostics: [{ code: 'VALIDATION_REJECTED', ruleId: 'amount' }],
        }),
      }}
      confirmValidationWarning={confirmValidationWarning}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: '12' } });
  fireEvent.keyDown(editor, { key: 'Enter' });

  await waitFor(() => expect(confirmValidationWarning).toHaveBeenCalledTimes(1));
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());
  expect(legacyProjection(ref.current!.getDocument())[0]?.rows?.['0']).toMatchObject({
    cells: { 0: { text: '12' } },
  });
});

it('keeps stale async editor validation from overwriting a newer revision', async () => {
  const ref = createRef<TegoSheetHandle>();
  const validation = deferredValidation();
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={validatedDocument()}
      validationEngine={validation.engine}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: '12' } });
  fireEvent.keyDown(editor, { key: 'Enter' });
  ref.current!.setCellText(
    { sheet: ref.current!.getDocument().workbook.sheets[0]!.id as never, row: 0, column: 0 },
    'newer',
  );
  validation.release({ status: 'accepted', diagnostics: [] });

  await waitFor(() => expect(rendered.getByRole('textbox', { name: /cell editor/i })).toBe(editor));
  expect(legacyProjection(ref.current!.getDocument())[0]?.rows?.['0']).toMatchObject({
    cells: { 0: { text: 'newer' } },
  });
});

it('cancels pending async editor validation when the component is disposed', async () => {
  const validation = deferredValidation();
  const onDocumentChange = vi.fn();
  const rendered = render(
    <TegoSheet
      defaultDocument={validatedDocument()}
      validationEngine={validation.engine}
      onDocumentChange={onDocumentChange}
    />,
  );
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: '12' } });
  fireEvent.keyDown(editor, { key: 'Enter' });
  rendered.unmount();
  validation.release({ status: 'accepted', diagnostics: [] });
  await Promise.resolve();
  expect(onDocumentChange).not.toHaveBeenCalled();
});

it('@parity:formulas.editor-display keeps typing local, commits once, creates one undo entry, and preserves callback order', async () => {
  const ref = createRef<TegoSheetHandle>();
  const order: string[] = [];
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={testDocument([{ rows: { 0: { cells: { 0: { text: 'old' } } } } }])}
      onDocumentChange={() => order.push('change')}
      onCellEdit={() => order.push('cell-edit')}
      onSelectionChange={() => order.push('selection')}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });

  fireEvent.change(editor, { target: { value: '=SUM(A1:A2)' } });
  const sheet = legacyProjection(ref.current!.getDocument())[0]!;
  expect(sheet.rows?.['0']).toMatchObject({ cells: { 0: { text: 'old' } } });
  expect(rendered.getByRole('listbox', { name: /formula suggestions/i })).toBeTruthy();

  fireEvent.keyDown(editor, { key: 'Enter' });
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());
  expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).toMatchObject({
    cells: { 0: { text: '=SUM(A1:A2)' } },
  });
  expect(order).toEqual(['change', 'cell-edit', 'selection']);

  order.length = 0;
  ref.current!.undo();
  expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).toMatchObject({
    cells: { 0: { text: 'old' } },
  });
  expect(order).toEqual(['change']);
  await waitFor(() =>
    expect(rendered.getByRole('button', { name: 'Undo' }).hasAttribute('disabled')).toBe(true),
  );
  const afterFirstUndo = legacyProjection(ref.current!.getDocument());
  const notifications = order.length;
  ref.current!.undo();
  expect(legacyProjection(ref.current!.getDocument())).toEqual(afterFirstUndo);
  expect(order).toHaveLength(notifications);
});

it('@parity:editing.inline-editor cancels on Escape and commits exactly once on blur', async () => {
  const ref = createRef<TegoSheetHandle>();
  const onChange = vi.fn();
  const rendered = render(
    <TegoSheet ref={ref} defaultDocument={testDocument([{}])} onDocumentChange={onChange} />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);

  fireEvent.keyDown(window, { key: 'x' });
  let editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'cancelled' } });
  fireEvent.keyDown(editor, { key: 'Escape' });
  expect(onChange).not.toHaveBeenCalled();

  fireEvent.keyDown(window, { key: 'y' });
  editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'saved' } });
  fireEvent.blur(editor);
  fireEvent.blur(editor);
  expect(onChange).toHaveBeenCalledOnce();
});

it('@parity:input.desktop-editing commits once on Tab and pointer navigation with selection before paint', async () => {
  const ref = createRef<TegoSheetHandle>();
  const order: string[] = [];
  const selections: number[] = [];
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={testDocument([{ rows: { len: 2 }, cols: { len: 3 } }])}
      onDocumentChange={() => order.push('change')}
      onCellEdit={() => order.push('cell-edit')}
      onSelectionChange={(selection) => {
        order.push('selection');
        selections.push(selection.active.column);
      }}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'a' });
  let editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'tabbed' } });
  fireEvent.keyDown(editor, { key: 'Tab' });
  expect(order).toEqual(['change', 'cell-edit', 'selection']);
  expect(selections).toEqual([1]);

  order.length = 0;
  selections.length = 0;
  fireEvent.keyDown(window, { key: 'b' });
  editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'pointed' } });
  fireEvent.pointerDown(root, { button: 0, buttons: 1, clientX: 260, clientY: 40 });
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());

  expect(order).toEqual(['change', 'cell-edit', 'selection']);
  expect(selections).toEqual([2]);
  expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).toMatchObject({
    cells: { 1: { text: 'pointed' } },
  });
});

it('preserves selection, scroll, and active editing across controlled acknowledgement', async () => {
  const ref = createRef<TegoSheetHandle>();
  const value: WorkbookInput = [{ name: 'Controlled', rows: { len: 5 }, cols: { len: 20 } }];
  let sheet: Parameters<TegoSheetHandle['setCellText']>[0]['sheet'] | undefined;
  let checkpoint: SpreadsheetDocument | undefined;
  let selectedColumn = -1;
  const onChange = (next: SpreadsheetDocument) => {
    checkpoint = next;
  };
  const onSelectionChange: NonNullable<Parameters<typeof TegoSheet>[0]['onSelectionChange']> = (
    next,
  ) => {
    sheet = next.sheet;
    selectedColumn = next.active.column;
  };
  const rendered = render(
    <TegoSheet
      ref={ref}
      document={testDocument(value)}
      onDocumentChange={onChange}
      onSelectionChange={onSelectionChange}
    />,
  );
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root, 300, 200);
  fireEvent.focusIn(root);
  for (let column = 0; column < 8; column += 1) fireEvent.keyDown(window, { key: 'ArrowRight' });
  expect(sheet).toBeDefined();
  expect(selectedColumn).toBe(8);
  ref.current!.setCellText({ sheet: sheet!, row: 0, column: 0 }, 'accepted');
  expect(checkpoint).toBeDefined();
  fireEvent.keyDown(window, { key: 'x' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'draft' } });
  const editorHost = editor.closest<HTMLElement>('.tego-sheet__editor')!;
  const preScrollLeft = editorHost.style.left;
  fireEvent.wheel(root, { deltaX: 200, deltaY: 0 });
  await waitFor(() => expect(editorHost.style.left).not.toBe(preScrollLeft));
  const scrolledLeft = editorHost.style.left;

  const acknowledgement = structuredClone(checkpoint!);
  rendered.rerender(
    <TegoSheet
      ref={ref}
      document={acknowledgement}
      onDocumentChange={onChange}
      onSelectionChange={onSelectionChange}
    />,
  );
  const currentEditor = rendered.getByRole('textbox', {
    name: /cell editor/i,
  }) as HTMLTextAreaElement;
  const currentEditorHost = currentEditor.closest<HTMLElement>('.tego-sheet__editor')!;
  expect(currentEditor).toBe(editor);
  expect(currentEditor.value).toBe('draft');
  expect(currentEditorHost.style.left).toBe(scrolledLeft);
  expect(document.activeElement).toBe(currentEditor);
  expect(selectedColumn).toBe(8);

  fireEvent.keyDown(currentEditor, { key: 'Escape' });
  fireEvent.keyDown(window, { key: 'ArrowRight' });
  expect(selectedColumn).toBe(9);
  fireEvent.keyDown(window, { key: 'x' });
  expect(await rendered.findByRole('textbox', { name: /cell editor/i })).toBeTruthy();

  rendered.rerender(<TegoSheet ref={ref} document={testDocument([{ name: 'Replacement' }])} />);
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());

  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'y' });
  expect(await rendered.findByRole('textbox', { name: /cell editor/i })).toBeTruthy();
  rendered.rerender(
    <TegoSheet ref={ref} document={testDocument([{ name: 'Replacement' }])} readOnly />,
  );
  await waitFor(() => expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull());
});

it('keeps the editor anchored to engine geometry through scroll and resize', async () => {
  const rendered = render(
    <TegoSheet defaultDocument={testDocument([{ rows: { len: 5 }, cols: { len: 10 } }])} />,
  );
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root, 300, 200);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'ArrowRight' });
  fireEvent.keyDown(window, { key: 'ArrowRight' });
  fireEvent.keyDown(window, { key: 'F2' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  const host = editor.closest<HTMLElement>('.tego-sheet__editor')!;
  const initialLeft = host.style.left;

  fireEvent.wheel(root, { deltaX: 100, deltaY: 0 });
  await waitFor(() => expect(host.style.left).not.toBe(initialLeft));
  const scrolledLeft = host.style.left;

  sizeRoot(root, 2_000, 1_000);
  await waitFor(() => expect(host.style.left).not.toBe(scrolledLeft));
  expect(rendered.getByRole('textbox', { name: /cell editor/i })).toBe(editor);
});

it('does not apply a stale Tab selection when onChange synchronously replaces controlled data', async () => {
  const selections = vi.fn();
  function Host() {
    const [value, setValue] = useState<WorkbookInput>([{ name: 'Before' }]);
    return (
      <TegoSheet
        document={testDocument(value)}
        onDocumentChange={() => flushSync(() => setValue([{ name: 'Replacement' }]))}
        onSelectionChange={selections}
      />
    );
  }
  const rendered = render(<Host />);
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'ArrowRight' });
  fireEvent.keyDown(window, { key: 'ArrowLeft' });
  selections.mockClear();
  fireEvent.keyDown(window, { key: 'x' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'draft' } });

  fireEvent.keyDown(editor, { key: 'Tab' });

  await waitFor(() => expect(rendered.getByRole('tab', { name: 'Replacement' })).toBeTruthy());
  expect(selections).not.toHaveBeenCalled();
  expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull();
});

it('does not finish stale Enter selection work after onChange unmounts the sheet', async () => {
  const selections = vi.fn();
  const rendered = render(
    <TegoSheet
      defaultDocument={testDocument([{}])}
      onDocumentChange={() => rendered.unmount()}
      onSelectionChange={selections}
    />,
  );
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'ArrowRight' });
  fireEvent.keyDown(window, { key: 'ArrowLeft' });
  selections.mockClear();
  fireEvent.keyDown(window, { key: 'x' });
  const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
  fireEvent.change(editor, { target: { value: 'draft' } });

  expect(() => fireEvent.keyDown(editor, { key: 'Enter' })).not.toThrow();
  expect(selections).not.toHaveBeenCalled();
  expect(rendered.container.childElementCount).toBe(0);
});

it.each(['onDocumentChange', 'onCellEdit'] as const)(
  'closes the editor and preserves the original %s exception after commit',
  async (callback) => {
    const ref = createRef<TegoSheetHandle>();
    const consumerError = new Error(`${callback} failed`);
    const callbacks =
      callback === 'onDocumentChange'
        ? {
            onDocumentChange: () => {
              throw consumerError;
            },
          }
        : {
            onCellEdit: () => {
              throw consumerError;
            },
          };
    const rendered = render(
      <TegoSheet ref={ref} defaultDocument={testDocument([{}])} {...callbacks} />,
    );
    await waitFor(() => expect(ref.current).not.toBeNull());
    const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
    sizeRoot(root);
    root.focus();
    fireEvent.keyDown(root, { key: 'x' });
    const editor = await rendered.findByRole('textbox', { name: /cell editor/i });
    fireEvent.change(editor, { target: { value: 'committed before callback failure' } });
    const reported: unknown[] = [];
    const capture = (event: ErrorEvent) => {
      reported.push(event.error);
      event.preventDefault();
    };
    window.addEventListener('error', capture);

    try {
      fireEvent.keyDown(editor, { key: 'Enter' });
    } finally {
      window.removeEventListener('error', capture);
    }

    expect(reported).toEqual([consumerError]);
    expect(rendered.queryByRole('textbox', { name: /cell editor/i })).toBeNull();
    expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).toMatchObject({
      cells: { 0: { text: 'committed before callback failure' } },
    });
    expect(() => rendered.unmount()).not.toThrow();
  },
);

it('clears with Delete without reporting a paste event', async () => {
  const ref = createRef<TegoSheetHandle>();
  const onPaste = vi.fn();
  const changes: string[] = [];
  const rendered = render(
    <TegoSheet
      ref={ref}
      defaultDocument={testDocument([{ rows: { 0: { cells: { 0: { text: 'erase me' } } } } }])}
      onDocumentChange={(_value, change) => changes.push(change.kind)}
      onPaste={onPaste}
    />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  fireEvent.focusIn(root);
  fireEvent.keyDown(window, { key: 'Delete' });

  expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).toMatchObject({
    cells: { 0: {} },
  });
  expect(legacyProjection(ref.current!.getDocument())[0]!.rows?.['0']).not.toHaveProperty(
    'cells.0.text',
  );
  expect(changes).toEqual(['cell']);
  expect(onPaste).not.toHaveBeenCalled();
});

it('does not notify or repaint React selection for an identical engine selection', () => {
  const onSelectionChange = vi.fn();
  const rendered = render(
    <TegoSheet defaultDocument={testDocument([{}])} onSelectionChange={onSelectionChange} />,
  );
  const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
  sizeRoot(root);
  fireEvent.focusIn(root);

  fireEvent.keyDown(window, { key: 'ArrowLeft' });
  fireEvent.pointerDown(root, { button: 0, buttons: 1, clientX: 70, clientY: 40 });

  expect(onSelectionChange).not.toHaveBeenCalled();
});
