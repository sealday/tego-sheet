import { expect, it, vi } from 'vitest';
import { TegoSheetException, type WorkbookInput } from '../../src/core';
import { renderSheet } from '../helpers/render-sheet';
import { testDocument } from '../helpers/workbook-builders';

it('throws a synchronous contract exception for mixed control modes', () => {
  vi.spyOn(console, 'error').mockImplementation(() => undefined);

  expect(() =>
    renderSheet({
      document: testDocument([]),
      defaultDocument: testDocument([]),
    } as never),
  ).toThrowError(
    expect.objectContaining({
      code: 'INVALID_COMMAND',
      recoverable: false,
    }),
  );
});

it('throws when a mounted boundary switches from controlled to uncontrolled', () => {
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
  const rendered = renderSheet({ document: testDocument([]) });

  expect(() => rendered.rerenderProps({ defaultDocument: testDocument([]) })).toThrow(
    TegoSheetException,
  );
});

it('throws when a mounted boundary switches from uncontrolled to controlled', () => {
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
  const rendered = renderSheet({ defaultDocument: testDocument([]) });

  expect(() => rendered.rerenderProps({ document: testDocument([]) })).toThrow(TegoSheetException);
});

it('throws invalid initial workbook data during initialization', () => {
  vi.spyOn(console, 'error').mockImplementation(() => undefined);

  expect(() => renderSheet({ defaultDocument: testDocument({ rows: { len: -1 } }) })).toThrowError(
    expect.objectContaining({ code: 'INVALID_DATA', recoverable: false }),
  );
});

it('does not reinterpret explicit null as a missing workbook prop', () => {
  vi.spyOn(console, 'error').mockImplementation(() => undefined);

  expect(() =>
    renderSheet({
      defaultDocument: testDocument(null as unknown as WorkbookInput),
    }),
  ).toThrowError(expect.objectContaining({ code: 'INVALID_DATA' }));
});

it('applies live readOnly before a parent layout effect can dispatch', () => {
  let shouldDispatch = false;
  let parentOutcome: ReturnType<typeof rendered.runtime.dispatchUi> | undefined;
  const rendered = renderSheet(
    { defaultDocument: testDocument([{}]), readOnly: false },
    {
      onParentLayout(runtime) {
        if (!shouldDispatch || runtime === null) return;
        parentOutcome = runtime.dispatchUi(
          {
            type: 'set-cell-text',
            address: { sheet, row: 0, column: 0 },
            text: 'must be rejected',
          },
          'keyboard',
        );
      },
    },
  );
  const sheet = rendered.runtime.epoch.snapshot.sheets[0]!.id;

  shouldDispatch = true;
  rendered.rerenderProps({ defaultDocument: testDocument([{}]), readOnly: true });
  shouldDispatch = false;

  expect(parentOutcome?.status).toBe('rejected');
  expect(rendered.runtime.epoch.controller.getCellText({ sheet, row: 0, column: 0 })).toBe('');
});
