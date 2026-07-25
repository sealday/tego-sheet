import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { IsolatedBrowserPrintAdapter, type GeneratedDocument } from '../../src';
import { downloadBlob } from '../../website/src/components/playground/output-download';
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
const outputStudioSource = readFileSync(
  resolve(process.cwd(), 'website/src/components/playground/output-studio.tsx'),
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

function createAdapterDoubles() {
  return {
    print: {
      print: vi.fn().mockResolvedValue({ pageCount: 2 }),
      dispose: vi.fn(),
    },
    pdf: {
      render: vi.fn().mockResolvedValue(new Blob(['pdf'], { type: 'application/pdf' })),
    },
    image: {
      render: vi.fn().mockResolvedValue([new Blob(['png'], { type: 'image/png' })] as const),
    },
    xlsx: {
      render: vi.fn().mockResolvedValue(
        new Blob(['xlsx'], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      ),
    },
  };
}

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
  vi.stubGlobal('URL', {
    createObjectURL: vi.fn(() => 'blob:output-studio'),
    revokeObjectURL: vi.fn(),
  });
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
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

it('passes the current generated document to every output adapter', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const adapters = createAdapterDoubles();
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Print 2 pages' }));
  await waitFor(() => expect(adapters.print.print).toHaveBeenCalledOnce());

  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download PNG page 1' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download XLSX' }));

  await waitFor(() => expect(adapters.xlsx.render).toHaveBeenCalledOnce());
  const generated = adapters.print.print.mock.calls[0]![0];
  expect(adapters.pdf.render.mock.calls[0]![0]).toBe(generated);
  expect(adapters.image.render.mock.calls[0]![0]).toBe(generated);
  expect(adapters.xlsx.render.mock.calls[0]![0]).toBe(generated);
});

it('uses the approved adapter options and deterministic Blob downloads', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const adapters = createAdapterDoubles();
  const downloads: string[] = [];
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(
    function (this: HTMLAnchorElement) {
      downloads.push(this.download);
    },
  );

  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  await screen.findByText('PDF downloaded');
  fireEvent.click(screen.getByRole('button', { name: 'Download PNG page 1' }));
  await screen.findByText('PNG page 1 downloaded');
  fireEvent.click(screen.getByRole('button', { name: 'Download XLSX' }));
  await screen.findByText('XLSX downloaded');

  expect(adapters.pdf.render).toHaveBeenCalledWith(generatedDocument, {
    pages: 'all',
    metadata: { title: 'Customer invoice' },
    tagged: false,
  });
  expect(adapters.image.render).toHaveBeenCalledWith(generatedDocument, {
    format: 'png',
    pages: [0],
    background: '#ffffff',
    dpi: 144,
  });
  expect(adapters.xlsx.render).toHaveBeenCalledWith(generatedDocument, {
    formulaMode: 'formula-and-cached-value',
    compatibility: 'excel',
  });
  const blobs = vi.mocked(URL.createObjectURL).mock.calls.map(([blob]) => blob as Blob);
  expect(blobs.map(({ type }) => type)).toEqual([
    'application/pdf',
    'image/png',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ]);
  expect(downloads).toEqual([
    'invoice-INV-2026-042.pdf',
    'invoice-INV-2026-042-page-1.png',
    'invoice-INV-2026-042.xlsx',
  ]);
  await waitFor(() => expect(URL.revokeObjectURL).toHaveBeenCalledTimes(3));
  expect(URL.revokeObjectURL).toHaveBeenNthCalledWith(1, 'blob:output-studio');
});

it('keeps other output actions available while one adapter is busy', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const pdf = deferred<Blob>();
  const adapters = createAdapterDoubles();
  adapters.pdf.render.mockReturnValue(pdf.promise);

  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));

  expect((screen.getByRole('button', { name: 'Download PDF' }) as HTMLButtonElement).disabled).toBe(
    true,
  );
  expect(
    (screen.getByRole('button', { name: 'Download PNG page 1' }) as HTMLButtonElement).disabled,
  ).toBe(false);
  expect(
    (screen.getByRole('button', { name: 'Download XLSX' }) as HTMLButtonElement).disabled,
  ).toBe(false);

  pdf.resolve(new Blob(['pdf'], { type: 'application/pdf' }));
  await screen.findByText('PDF downloaded');
});

it('tracks concurrent output actions as independently busy', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const pdf = deferred<Blob>();
  const png = deferred<readonly Blob[]>();
  const adapters = createAdapterDoubles();
  adapters.pdf.render.mockReturnValue(pdf.promise);
  adapters.image.render.mockReturnValue(png.promise);

  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download PNG page 1' }));

  expect((screen.getByRole('button', { name: 'Download PDF' }) as HTMLButtonElement).disabled).toBe(
    true,
  );
  expect(
    (screen.getByRole('button', { name: 'Download PNG page 1' }) as HTMLButtonElement).disabled,
  ).toBe(true);
  expect(
    (screen.getByRole('button', { name: 'Download XLSX' }) as HTMLButtonElement).disabled,
  ).toBe(false);

  pdf.resolve(new Blob(['pdf'], { type: 'application/pdf' }));
  png.resolve([new Blob(['png'], { type: 'image/png' })]);
  await waitFor(() => expect(URL.createObjectURL).toHaveBeenCalledTimes(2));
});

it('removes the temporary anchor and revokes its Blob URL when clicking fails', async () => {
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {
    throw new Error('download blocked');
  });
  const blob = new Blob(['pdf'], { type: 'application/pdf' });

  expect(() => downloadBlob(blob, 'invoice.pdf')).toThrow('download blocked');
  expect(document.querySelector('a[download="invoice.pdf"]')).toBeNull();
  expect(URL.createObjectURL).toHaveBeenCalledWith(blob);
  await waitFor(() => expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:output-studio'));
});

it('announces PRINT_BLOCKED without discarding the generated preview', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const adapters = createAdapterDoubles();
  adapters.print.print.mockRejectedValue(
    Object.assign(new Error('Browser print was blocked by the host'), {
      code: 'PRINT_BLOCKED',
    }),
  );

  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Print 2 pages' }));

  expect((await screen.findByRole('alert')).textContent).toBe(
    'PRINT failed (PRINT_BLOCKED): Browser print was blocked by the host',
  );
  expect(screen.getAllByRole('article', { name: /Print page/u })).toHaveLength(2);
});

it('announces an XLSX rejection and leaves the preview visible', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const adapters = createAdapterDoubles();
  adapters.xlsx.render.mockRejectedValue(new Error('Workbook packaging failed'));

  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download XLSX' }));

  expect((await screen.findByRole('alert')).textContent).toBe(
    'XLSX failed: Workbook packaging failed',
  );
  expect(screen.getAllByRole('article', { name: /Print page/u })).toHaveLength(2);
  expect(
    (screen.getByRole('button', { name: 'Download XLSX' }) as HTMLButtonElement).disabled,
  ).toBe(false);
});

it('disposes the default browser print adapter on unmount', () => {
  pipeline.renderOutputRevision.mockReturnValue(new Promise(() => undefined));
  const dispose = vi.spyOn(IsolatedBrowserPrintAdapter.prototype, 'dispose');

  const rendered = render(<OutputStudio />);
  rendered.unmount();

  expect(dispose).toHaveBeenCalledOnce();
});

it('keeps every output disabled when generation is blocked', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [
      {
        code: 'MISSING_DATA',
        severity: 'error',
        domain: 'template',
        stage: 'render',
        message: 'Invoice data is missing',
      },
    ],
  });

  render(<OutputStudio adapters={createAdapterDoubles()} />);
  await screen.findByText('Generation is blocked. Review the diagnostics.');

  for (const name of ['Print 0 pages', 'Download PDF', 'Download PNG page 1', 'Download XLSX']) {
    expect((screen.getByRole('button', { name }) as HTMLButtonElement).disabled).toBe(true);
  }
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
  for (const name of ['Download PDF', 'Download PNG page 1', 'Download XLSX']) {
    expect((screen.getByRole('button', { name }) as HTMLButtonElement).disabled).toBe(true);
  }
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
  expect(outputStudioSource).toContain("from 'tego-sheet/output/pdf'");
  expect(outputStudioSource).toContain("from 'tego-sheet/output/image'");
  expect(outputStudioSource).toContain("from 'tego-sheet/output/xlsx'");
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
  expect(playgroundStyles).toMatch(
    /\.templateSheet input:not\(\[type=['"]checkbox['"]\]\):not\(\[type=['"]radio['"]\]\)\s*{[^}]*min-height:\s*2\.75rem/s,
  );
  expect(playgroundStyles).toMatch(
    /\.templateSheet label:has\(> input\[type=['"]checkbox['"]\]\),\s*\.templateSheet label:has\(> input\[type=['"]radio['"]\]\)\s*{[^}]*min-height:\s*2\.75rem/s,
  );
  expect(playgroundStyles).not.toMatch(/\.templateSheet input\s*{[^}]*min-height:/s);
  expect(playgroundStyles).toMatch(/\.templateSheet input:focus-visible\s*{[^}]*outline:/s);
});
