import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode } from 'react';
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
  readonly reject: (reason: unknown) => void;
} {
  let resolvePromise!: (value: Value) => void;
  let rejectPromise!: (reason: unknown) => void;
  const promise = new Promise<Value>((resolve, reject) => {
    resolvePromise = resolve;
    rejectPromise = reject;
  });
  return { promise, resolve: resolvePromise, reject: rejectPromise };
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
  expect(adapters.print.print).toHaveBeenCalledWith(generatedDocument, {
    signal: expect.any(AbortSignal),
  });

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
    signal: expect.any(AbortSignal),
  });
  expect(adapters.image.render).toHaveBeenCalledWith(generatedDocument, {
    format: 'png',
    pages: [0],
    background: '#ffffff',
    dpi: 144,
    signal: expect.any(AbortSignal),
  });
  expect(adapters.xlsx.render).toHaveBeenCalledWith(generatedDocument, {
    formulaMode: 'formula-and-cached-value',
    compatibility: 'excel',
    signal: expect.any(AbortSignal),
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

it('recreates an owned default print adapter after the StrictMode effect probe', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const disposed = new WeakSet<object>();
  vi.spyOn(IsolatedBrowserPrintAdapter.prototype, 'dispose').mockImplementation(
    function (this: IsolatedBrowserPrintAdapter) {
      disposed.add(this);
    },
  );
  const print = vi
    .spyOn(IsolatedBrowserPrintAdapter.prototype, 'print')
    .mockImplementation(function (this: IsolatedBrowserPrintAdapter) {
      return disposed.has(this)
        ? Promise.reject(new Error('print adapter is disposed'))
        : Promise.resolve({
            pageIds: ['invoice-page-1', 'invoice-page-2'],
            pageCount: 2,
            cleanupReason: 'afterprint' as const,
          });
    });

  render(
    <StrictMode>
      <OutputStudio />
    </StrictMode>,
  );
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Print 2 pages' }));

  await screen.findByText('Print dialog opened');
  expect(print).toHaveBeenCalledOnce();
  expect(disposed.has(print.mock.instances[0]!)).toBe(false);
});

it('switches adapter injections without disposing injected adapters or retaining an owned default', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const ownedDispose = vi.spyOn(IsolatedBrowserPrintAdapter.prototype, 'dispose');
  const firstInjected = createAdapterDoubles();
  const secondInjected = createAdapterDoubles();
  const rendered = render(<OutputStudio />);
  await screen.findByText('GeneratedDocument · revision 1');

  rendered.rerender(<OutputStudio adapters={firstInjected} />);
  rendered.rerender(<OutputStudio adapters={secondInjected} />);
  fireEvent.click(screen.getByRole('button', { name: 'Print 2 pages' }));
  await waitFor(() => expect(secondInjected.print.print).toHaveBeenCalledOnce());

  expect(firstInjected.print.print).not.toHaveBeenCalled();
  expect(firstInjected.print.dispose).not.toHaveBeenCalled();
  expect(secondInjected.print.dispose).not.toHaveBeenCalled();
  expect(ownedDispose).toHaveBeenCalledOnce();
  rendered.unmount();
  expect(secondInjected.print.dispose).not.toHaveBeenCalled();
});

it('never disposes an injected adapter when switching back to an owned default', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const injected = createAdapterDoubles();
  const rendered = render(<OutputStudio adapters={injected} />);
  await screen.findByText('GeneratedDocument · revision 1');

  rendered.rerender(<OutputStudio />);
  rendered.unmount();

  expect(injected.print.dispose).not.toHaveBeenCalled();
});

it('cancels old requests and clears their busy state when injected adapters change', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const pending = deferred<Blob>();
  const firstInjected = createAdapterDoubles();
  firstInjected.pdf.render.mockReturnValue(pending.promise);
  const secondInjected = createAdapterDoubles();
  const rendered = render(<OutputStudio adapters={firstInjected} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  const signal = firstInjected.pdf.render.mock.calls[0]![1].signal as AbortSignal;

  rendered.rerender(<OutputStudio adapters={secondInjected} />);

  expect(signal.aborted).toBe(true);
  await waitFor(() =>
    expect(
      (screen.getByRole('button', { name: 'Download PDF' }) as HTMLButtonElement).disabled,
    ).toBe(false),
  );
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  await waitFor(() => expect(secondInjected.pdf.render).toHaveBeenCalledOnce());
});

it('aborts a pending output and ignores its late completion after a draft change', async () => {
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
  const pdfButton = screen.getByRole('button', { name: 'Download PDF' });
  expect(pdfButton.getAttribute('aria-busy')).toBe('true');
  expect(screen.getByText('Generating PDF…')).toBeTruthy();
  const signal = adapters.pdf.render.mock.calls[0]![1].signal as AbortSignal;
  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(screen.getByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });

  expect(signal.aborted).toBe(true);
  expect(screen.getByText('PDF generation cancelled.')).toBeTruthy();
  expect(pdfButton.getAttribute('aria-busy')).toBe('false');
  pdf.resolve(new Blob(['late-pdf'], { type: 'application/pdf' }));
  await act(async () => pdf.promise);
  expect(URL.createObjectURL).not.toHaveBeenCalled();
  expect(screen.queryByText('PDF downloaded')).toBeNull();
});

it('clears completed output outcomes when the draft moves beyond their generated revision', async () => {
  pipeline.renderOutputRevision.mockImplementation(async ({ revision }) => ({
    revision,
    diagnostics: [],
    document: generatedDocument,
  }));
  const adapters = createAdapterDoubles();
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  await screen.findByText('PDF downloaded');

  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(screen.getByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });
  expect(screen.queryByText('PDF downloaded')).toBeNull();

  fireEvent.click(screen.getByRole('button', { name: 'Apply & regenerate' }));
  await screen.findByText('GeneratedDocument · revision 2');
  expect(screen.queryByText('PDF downloaded')).toBeNull();
});

it('ignores an old revision output after a regenerated revision starts a new request', async () => {
  pipeline.renderOutputRevision.mockImplementation(async ({ revision }) => ({
    revision,
    diagnostics: [],
    document: generatedDocument,
  }));
  const oldPdf = deferred<Blob>();
  const newPdf = deferred<Blob>();
  const adapters = createAdapterDoubles();
  adapters.pdf.render.mockReturnValueOnce(oldPdf.promise).mockReturnValueOnce(newPdf.promise);
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  const oldSignal = adapters.pdf.render.mock.calls[0]![1].signal as AbortSignal;

  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  const data = JSON.parse((screen.getByLabelText('Data JSON') as HTMLTextAreaElement).value);
  data.invoice.id = 'INV-REVISION-2';
  fireEvent.change(screen.getByLabelText('Data JSON'), {
    target: { value: JSON.stringify(data, null, 2) },
  });
  fireEvent.click(screen.getByRole('button', { name: 'Apply & regenerate' }));
  await screen.findByText('GeneratedDocument · revision 2');
  expect(oldSignal.aborted).toBe(true);

  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  newPdf.resolve(new Blob(['new'], { type: 'application/pdf' }));
  await screen.findByText('PDF downloaded');
  oldPdf.resolve(new Blob(['old'], { type: 'application/pdf' }));
  await act(async () => oldPdf.promise);

  expect(URL.createObjectURL).toHaveBeenCalledOnce();
  expect(URL.createObjectURL).toHaveBeenCalledWith(
    expect.objectContaining({ type: 'application/pdf' }),
  );
});

it('aborts pending output and prevents downloads after unmount', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const png = deferred<readonly Blob[]>();
  const adapters = createAdapterDoubles();
  adapters.image.render.mockReturnValue(png.promise);
  const rendered = render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Download PNG page 1' }));
  const signal = adapters.image.render.mock.calls[0]![1].signal as AbortSignal;
  rendered.unmount();

  expect(signal.aborted).toBe(true);
  png.resolve([new Blob(['late-png'], { type: 'image/png' })]);
  await act(async () => png.promise);
  expect(URL.createObjectURL).not.toHaveBeenCalled();
});

it('retains independent output outcomes when overlapping requests settle out of order', async () => {
  pipeline.renderOutputRevision.mockResolvedValue({
    revision: 1,
    diagnostics: [],
    document: generatedDocument,
  });
  const pdf = deferred<Blob>();
  const xlsx = deferred<Blob>();
  const adapters = createAdapterDoubles();
  adapters.pdf.render.mockReturnValue(pdf.promise);
  adapters.xlsx.render.mockReturnValue(xlsx.promise);
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download XLSX' }));
  xlsx.reject(new Error('Workbook packaging failed'));
  const xlsxAlert = await screen.findByRole('alert');
  expect(xlsxAlert.textContent).toBe('XLSX failed: Workbook packaging failed');

  pdf.resolve(new Blob(['pdf'], { type: 'application/pdf' }));
  await screen.findByText('PDF downloaded');
  expect(screen.getByRole('alert')).toBe(xlsxAlert);
  expect(screen.getByText('PDF downloaded')).toBeTruthy();
});

it('uses metadata committed with the regenerated document for filenames', async () => {
  pipeline.renderOutputRevision.mockImplementation(async ({ revision }) => ({
    revision,
    diagnostics: [],
    document: generatedDocument,
  }));
  const adapters = createAdapterDoubles();
  const downloads: string[] = [];
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(
    function (this: HTMLAnchorElement) {
      downloads.push(this.download);
    },
  );
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  const data = JSON.parse((screen.getByLabelText('Data JSON') as HTMLTextAreaElement).value);
  data.invoice.id = 'INV-REGENERATED-7';
  fireEvent.change(screen.getByLabelText('Data JSON'), {
    target: { value: JSON.stringify(data, null, 2) },
  });
  fireEvent.click(screen.getByRole('button', { name: 'Apply & regenerate' }));
  await screen.findByText('GeneratedDocument · revision 2');
  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  await screen.findByText('PDF downloaded');

  expect(downloads).toEqual(['invoice-INV-REGENERATED-7.pdf']);
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
    /@media \(max-width: 42rem\)[\s\S]*?\.outputStudioGrid\s*{[^}]*display:\s*flex;[^}]*flex-direction:\s*column/,
  );
  expect(playgroundStyles).toMatch(
    /@media \(max-width: 42rem\)[\s\S]*?\.exactPreview\s*{[^}]*order:\s*1[^}]*}[\s\S]*?\.outputInputs\s*{[^}]*order:\s*2[^}]*}[\s\S]*?\.pipelineOutputs\s*{[^}]*order:\s*3/,
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
