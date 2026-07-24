import { act, cleanup, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet, type TegoSheetHandle } from '../../../src';
import { createPersistenceSession } from '../../../src/integrations/persistence';
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

it('binds committed UI/ref transactions and exposes live persistence status', async () => {
  const document = testDocument([{ name: 'A' }]);
  const ref = createRef<TegoSheetHandle>();
  const save = vi.fn(async (request) => ({
    status: 'saved' as const,
    revision: 'revision-2',
    persistedTransactionIds: request.transactions.map(({ id }: { id: string }) => id),
  }));
  const session = createPersistenceSession({
    documentId: document.id,
    initialRevision: 'revision-1',
    adapter: { save },
    requestId: () => 'request-1',
    autosaveDelayMs: 60_000,
  });
  const rendered = render(
    <TegoSheet ref={ref} defaultDocument={document} persistenceSession={session} />,
  );
  await waitFor(() => expect(ref.current).not.toBeNull());
  const sheet = document.workbook.sheets[0]!.id as never;

  act(() => ref.current!.setCellText({ sheet, row: 0, column: 0 }, 'tracked'));
  await waitFor(() =>
    expect(rendered.container.firstElementChild?.getAttribute('data-persistence-status')).toBe(
      'dirty',
    ),
  );
  await act(async () => {
    await session.save();
  });
  expect(save).toHaveBeenCalledTimes(1);
  expect(rendered.container.firstElementChild?.getAttribute('data-persistence-status')).toBe(
    'clean',
  );

  session.dispose();
});
