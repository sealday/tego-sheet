import { cleanup, render } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet } from '../../../src';
import { createCommentStore } from '../../../src/integrations/comments';
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

it('projects live semantic comment markers and an explicit print policy', () => {
  let revision = 0;
  const store = createCommentStore({
    documentId: 'document-1',
    actorId: 'actor-1',
    permissions: () => ({ revision: 'p1', actorId: 'actor-1', can: () => true }),
    nextId: () => 'thread-1',
    nextRevision: () => `thread-revision-${++revision}`,
  });
  store.create({
    anchor: { type: 'cell', cell: { sheetId: 'sheet-1', row: 0, column: 0 } },
    content: [{ text: 'Review this' }],
    expectedDocumentRevision: 'revision-1',
    currentDocumentRevision: 'revision-1',
  });
  const rendered = render(
    <TegoSheet
      defaultDocument={testDocument([{ name: 'A' }])}
      commentStore={store}
      commentPrintPolicy="markers"
    />,
  );

  expect(rendered.container.querySelector('[data-comment-marker="thread-1"]')?.textContent).toBe(
    '1',
  );
  expect(
    rendered.container
      .querySelector('[data-tego-sheet]')
      ?.getAttribute('data-comment-print-policy'),
  ).toBe('markers');
});
