import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import type { GeneratedDocument } from '../../src/template';
import { OutputStudio } from '../../website/src/components/playground/output-studio';
import { createCanvasHarness } from '../helpers/canvas-harness';

const pipeline = vi.hoisted(() => ({
  renderOutputRevision: vi.fn(),
}));

vi.mock('../../website/src/components/playground/output-studio-pipeline', () => pipeline);

const playgroundStyles = readFileSync(
  resolve(process.cwd(), 'website/src/components/playground/playground.module.css'),
  'utf8',
);

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

function deferred<Value>(): {
  readonly promise: Promise<Value>;
  readonly resolve: (value: Value) => void;
} {
  let resolvePromise!: (value: Value) => void;
  const promise = new Promise<Value>((resolve) => {
    resolvePromise = resolve;
  });
  return { promise, resolve: resolvePromise };
}

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
  vi.clearAllMocks();
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
  expect(
    screen.getByText('Preview, print, PDF, and PNG share one exact display list.'),
  ).toBeTruthy();
  expect(
    screen.getByText('XLSX uses the semantic workbook in the same GeneratedDocument.'),
  ).toBeTruthy();
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

  expect(
    screen.getByText('Preview is stale. Apply & regenerate to update every output.'),
  ).toBeTruthy();
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

it('keeps a newer draft stale when an older render resolves', async () => {
  const initialRender = deferred<{
    readonly revision: number;
    readonly diagnostics: readonly [];
    readonly document: GeneratedDocument;
  }>();
  pipeline.renderOutputRevision.mockReturnValue(initialRender.promise);

  render(<OutputStudio />);
  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(await screen.findByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });

  await act(async () => {
    initialRender.resolve({
      revision: 1,
      diagnostics: [],
      document: generatedDocument,
    });
    await initialRender.promise;
  });

  expect(
    screen.getByText('Preview is stale. Apply & regenerate to update every output.'),
  ).toBeTruthy();
  expect(
    (screen.getByRole('button', { name: 'Print 0 pages' }) as HTMLButtonElement).disabled,
  ).toBe(true);
});

it('aborts the regenerated request when the studio unmounts', async () => {
  const regeneratedRender = deferred<never>();
  pipeline.renderOutputRevision
    .mockResolvedValueOnce({
      revision: 1,
      diagnostics: [],
      document: generatedDocument,
    })
    .mockReturnValueOnce(regeneratedRender.promise);

  const rendered = render(<OutputStudio />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(screen.getByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });
  fireEvent.click(screen.getByRole('button', { name: 'Apply & regenerate' }));

  const regeneratedSignal = pipeline.renderOutputRevision.mock.calls[1]?.[0].signal as AbortSignal;
  expect(regeneratedSignal.aborted).toBe(false);
  rendered.unmount();
  expect(regeneratedSignal.aborted).toBe(true);
});

it('defines the approved responsive areas and accessible nested designer controls', () => {
  expect(playgroundStyles).toMatch(
    /\.outputStudioGrid\s*{[^}]*grid-template-areas:\s*['"]inputs preview outputs['"]/s,
  );
  expect(playgroundStyles).toMatch(/\.outputInputs\s*{[^}]*grid-area:\s*inputs/s);
  expect(playgroundStyles).toMatch(/\.exactPreview\s*{[^}]*grid-area:\s*preview/s);
  expect(playgroundStyles).toMatch(/\.pipelineOutputs\s*{[^}]*grid-area:\s*outputs/s);
  expect(playgroundStyles).toMatch(
    /@media \(max-width: 72rem\)[\s\S]*?grid-template-areas:\s*['"]preview preview['"]\s*['"]inputs outputs['"]/,
  );
  expect(playgroundStyles).toMatch(
    /@media \(max-width: 48rem\)[\s\S]*?grid-template-areas:\s*['"]preview['"]\s*['"]inputs['"]\s*['"]outputs['"]/,
  );
  expect(playgroundStyles).toMatch(/\.templateSheet input\s*{[^}]*min-height:\s*2\.75rem/s);
  expect(playgroundStyles).toMatch(/\.templateSheet input:focus-visible\s*{[^}]*outline:/s);
});
