import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import type { GeneratedDocument } from '../../src/template';
import { OutputStudio } from '../../website/src/components/playground/output-studio';
import { createCanvasHarness } from '../helpers/canvas-harness';

const pipeline = vi.hoisted(() => ({
  renderOutputRevision: vi.fn(),
}));

vi.mock('../../website/src/components/playground/output-studio-pipeline', () => pipeline);

const generatedDocument = {
  print: {
    pages: [
      {
        id: 'invoice-page-1',
        index: 0,
        width: 595,
        height: 842,
        targetId: 'invoice-sheet',
        rowStart: 0,
        rowEnd: 13,
      },
      {
        id: 'invoice-page-2',
        index: 1,
        width: 595,
        height: 842,
        targetId: 'invoice-sheet',
        rowStart: 14,
        rowEnd: 29,
      },
    ],
    displayList: {
      diagnostics: [],
      pages: [
        {
          index: 0,
          width: 595,
          height: 842,
          commands: [
            {
              kind: 'text',
              text: 'INVOICE',
              x: 24,
              y: 36,
              maxWidth: 180,
              fontFamily: 'Arial',
              fontSize: 18,
              color: '#111827',
              horizontalAlign: 'left',
            },
          ],
        },
        {
          index: 1,
          width: 595,
          height: 842,
          commands: [
            {
              kind: 'text',
              text: 'Thank you for your business.',
              x: 24,
              y: 36,
              maxWidth: 240,
              fontFamily: 'Arial',
              fontSize: 12,
              color: '#111827',
              horizontalAlign: 'left',
            },
          ],
        },
      ],
    },
  },
} as unknown as GeneratedDocument;

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

it('renders the prepared revision and explains the shared artifact', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });

  render(<OutputStudio />);

  expect(screen.getByRole('heading', { name: 'Output Studio' })).toBeTruthy();
  await waitFor(() =>
    expect(screen.getAllByRole('article', { name: /Print page/u })).toHaveLength(2),
  );
  expect(screen.getByText('GeneratedDocument · revision 1')).toBeTruthy();
  expect(screen.getByText(/One document · many outputs/u)).toBeTruthy();
});

it('keeps template edits as drafts until Apply and regenerate', async () => {
  pipeline.renderOutputRevision.mockImplementation(async ({ revision }) => ({
    revision,
    diagnostics: [],
    document: generatedDocument,
  }));

  render(<OutputStudio />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(screen.getByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });

  expect(screen.getByRole('status').textContent).toContain('Preview is stale');
  expect(
    (screen.getByRole('button', { name: 'Print 2 pages' }) as HTMLButtonElement).disabled,
  ).toBe(true);
  expect(document.querySelector('[data-mode="template"]')).toBeTruthy();

  fireEvent.click(screen.getByRole('button', { name: 'Apply & regenerate' }));
  await screen.findByText('GeneratedDocument · revision 2');
  expect(pipeline.renderOutputRevision).toHaveBeenLastCalledWith(
    expect.objectContaining({
      revision: 2,
      template: expect.objectContaining({
        bindings: expect.arrayContaining([
          expect.objectContaining({
            id: 'customer-name',
            expression: 'customer.legalName',
          }),
        ]),
      }),
    }),
  );
});
