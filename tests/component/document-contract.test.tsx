import { act, cleanup, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import {
  createSpreadsheetDocument,
  TegoSheet,
  type SpreadsheetDocument,
  type TegoSheetHandle,
  type TegoSheetProps,
} from '../../src';
import { sheetId } from '../../src/core';
import { createCanvasHarness } from '../helpers/canvas-harness';

afterEach(cleanup);
beforeEach(() => {
  const context = createCanvasHarness().canvas.getContext('2d');
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => context);
});

it('owns an uncontrolled defaultDocument and emits one frozen document per command', async () => {
  const input = createSpreadsheetDocument({
    id: 'document-1',
    sheetId: 'sheet-1',
    sheetName: 'Uncontrolled',
  });
  const onDocumentChange = vi.fn();
  const ref = createRef<TegoSheetHandle>();
  render(<TegoSheet ref={ref} defaultDocument={input} onDocumentChange={onDocumentChange} />);
  await waitFor(() => expect(ref.current).not.toBeNull());

  act(() => {
    ref.current?.setCellText({ sheet: sheetId('sheet-1'), row: 0, column: 0 }, 'committed');
  });

  expect(onDocumentChange).toHaveBeenCalledOnce();
  const emitted = onDocumentChange.mock.calls[0]?.[0] as SpreadsheetDocument;
  const snapshot = ref.current?.getDocument();
  expect(emitted).not.toBe(snapshot);
  expect(Object.isFrozen(emitted)).toBe(true);
  expect(Object.isFrozen(snapshot)).toBe(true);
  expect(snapshot?.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
    type: 'string',
    value: 'committed',
  });
});

it('reconciles a controlled document without sharing the caller reference', async () => {
  const initial = createSpreadsheetDocument({
    id: 'document-1',
    sheetId: 'sheet-1',
    sheetName: 'Initial',
  });
  const replacement = createSpreadsheetDocument({
    id: 'document-1',
    sheetId: 'sheet-1',
    sheetName: 'Replacement',
  });
  const ref = createRef<TegoSheetHandle>();
  const rendered = render(<TegoSheet ref={ref} document={initial} />);
  await waitFor(() => expect(ref.current).not.toBeNull());

  rendered.rerender(<TegoSheet ref={ref} document={replacement} />);

  expect(ref.current?.getDocument()).not.toBe(replacement);
  expect(ref.current?.getDocument().workbook.sheets[0]?.name).toBe('Replacement');
});

it('removes legacy mutable ingress from public props and handle types', () => {
  type HasLegacyProps =
    | ('value' extends keyof TegoSheetProps ? true : false)
    | ('defaultValue' extends keyof TegoSheetProps ? true : false)
    | ('onChange' extends keyof TegoSheetProps ? true : false);
  type HasLegacyHandle = 'getValue' extends keyof TegoSheetHandle ? true : false;
  const legacyProps: HasLegacyProps = false;
  const legacyHandle: HasLegacyHandle = false;
  expect(legacyProps).toBe(false);
  expect(legacyHandle).toBe(false);
});

it('makes controlled and uncontrolled document ownership mutually exclusive in props', () => {
  type AcceptsBoth = {
    document: SpreadsheetDocument;
    defaultDocument: SpreadsheetDocument;
  } extends TegoSheetProps
    ? true
    : false;
  type AcceptsNeither = Record<never, never> extends TegoSheetProps ? true : false;
  const acceptsBoth: AcceptsBoth = false;
  const acceptsNeither: AcceptsNeither = false;
  expect(acceptsBoth).toBe(false);
  expect(acceptsNeither).toBe(false);
});
