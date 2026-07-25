import {
  IsolatedBrowserPrintAdapter,
  TegoSheet,
  TemplatePreview,
  type Diagnostic,
  type GeneratedDocument,
} from 'tego-sheet';
import { ImageAdapter } from 'tego-sheet/output/image';
import { PdfAdapter } from 'tego-sheet/output/pdf';
import { XlsxAdapter } from 'tego-sheet/output/xlsx';
import 'tego-sheet/styles.css';
import {
  useCallback,
  useEffect,
  useReducer,
  useRef,
  useState,
  type ChangeEvent,
  type ReactElement,
} from 'react';
import { downloadBlob } from './output-download';
import { createInvoiceOutputFixture } from './output-studio-fixtures';
import {
  createOutputStudioState,
  outputFilename,
  reduceOutputStudioState,
  type OutputDocumentMetadata,
  type OutputKind,
} from './output-studio-model';
import { renderOutputRevision } from './output-studio-pipeline';
import styles from './playground.module.css';

export interface OutputStudioAdapters {
  readonly print: Pick<IsolatedBrowserPrintAdapter, 'print' | 'dispose'>;
  readonly pdf: Pick<PdfAdapter, 'render'>;
  readonly image: Pick<ImageAdapter, 'render'>;
  readonly xlsx: Pick<XlsxAdapter, 'render'>;
}

interface OutputStudioProps {
  readonly adapters?: OutputStudioAdapters;
}

function createOutputStudioAdapters(): OutputStudioAdapters {
  return {
    print: new IsolatedBrowserPrintAdapter(),
    pdf: new PdfAdapter(),
    image: new ImageAdapter(),
    xlsx: new XlsxAdapter(),
  };
}

function outputErrorMessage(kind: OutputKind, error: unknown): string {
  const name = kind === 'png' ? 'PNG' : kind.toUpperCase();
  const message = error instanceof Error ? error.message : 'Output generation failed';
  const code =
    typeof error === 'object' && error !== null && 'code' in error && typeof error.code === 'string'
      ? ` (${error.code})`
      : '';
  return `${name} failed${code}: ${message}`;
}

function outputDocumentMetadata(
  template: { readonly name: string },
  data: unknown,
): OutputDocumentMetadata {
  const invoice =
    typeof data === 'object' && data !== null && 'invoice' in data ? data.invoice : undefined;
  const invoiceId =
    typeof invoice === 'object' &&
    invoice !== null &&
    'id' in invoice &&
    typeof invoice.id === 'string'
      ? invoice.id
      : 'output';
  return { invoiceId, title: template.name };
}

interface ActiveOutputRequest {
  readonly requestId: number;
  readonly revision: number;
  readonly controller: AbortController;
}

export function OutputStudio({ adapters: injectedAdapters }: OutputStudioProps = {}): ReactElement {
  const [fixture] = useState(createInvoiceOutputFixture);
  const [draftTemplate, setDraftTemplate] = useState(fixture.template);
  const [draftData, setDraftData] = useState(() => JSON.stringify(fixture.data, null, 2));
  const [workbenchOpen, setWorkbenchOpen] = useState(false);
  const [dataError, setDataError] = useState('');
  const [state, dispatch] = useReducer(reduceOutputStudioState, undefined, createOutputStudioState);
  const revisionRef = useRef(1);
  const controllerRef = useRef<AbortController | null>(null);
  const adaptersRef = useRef<OutputStudioAdapters | null>(injectedAdapters ?? null);
  const adapterPropRef = useRef(injectedAdapters);
  const mountedRef = useRef(false);
  const generatedRevisionRef = useRef<number | null>(null);
  const outputRequestIdRef = useRef(0);
  const outputRequestsRef = useRef(new Map<OutputKind, ActiveOutputRequest>());

  const abortOutputRequests = useCallback((): void => {
    for (const request of outputRequestsRef.current.values()) request.controller.abort();
    outputRequestsRef.current.clear();
  }, []);

  const startRender = useCallback(
    (
      revision: number,
      template: typeof fixture.template,
      data: unknown,
      metadata: OutputDocumentMetadata,
    ): AbortController => {
      abortOutputRequests();
      generatedRevisionRef.current = null;
      controllerRef.current?.abort();
      const controller = new AbortController();
      controllerRef.current = controller;
      dispatch({ type: 'render-started', revision });
      void renderOutputRevision({
        revision,
        document: fixture.document,
        template,
        data,
        environment: fixture.environment,
        signal: controller.signal,
      })
        .then((result) => {
          if (controller.signal.aborted) return;
          if (result.document === undefined) {
            dispatch({
              type: 'render-blocked',
              revision,
              diagnostics: result.diagnostics,
            });
          } else {
            generatedRevisionRef.current = revision;
            dispatch({
              type: 'render-succeeded',
              revision,
              document: result.document,
              metadata,
              diagnostics: result.diagnostics,
            });
          }
        })
        .catch((error: unknown) => {
          if (controller.signal.aborted) return;
          const diagnostic: Diagnostic = {
            code: 'RENDER_FAILED',
            severity: 'error',
            domain: 'template',
            stage: 'render',
            message: error instanceof Error ? error.message : 'Unable to generate the document',
          };
          dispatch({ type: 'render-blocked', revision, diagnostics: [diagnostic] });
        });
      return controller;
    },
    [abortOutputRequests, fixture],
  );

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      generatedRevisionRef.current = null;
      controllerRef.current?.abort();
      abortOutputRequests();
    };
  }, [abortOutputRequests]);

  useEffect(() => {
    const adaptersChanged = adapterPropRef.current !== injectedAdapters;
    adapterPropRef.current = injectedAdapters;
    const adapters = injectedAdapters ?? createOutputStudioAdapters();
    adaptersRef.current = adapters;
    if (adaptersChanged) dispatch({ type: 'outputs-cancelled' });
    return () => {
      if (adaptersRef.current === adapters) adaptersRef.current = null;
      abortOutputRequests();
      if (injectedAdapters === undefined) adapters.print.dispose();
    };
  }, [abortOutputRequests, injectedAdapters]);

  useEffect(() => {
    startRender(
      1,
      fixture.template,
      fixture.data,
      outputDocumentMetadata(fixture.template, fixture.data),
    );
    return () => controllerRef.current?.abort();
  }, [fixture, startRender]);

  const markDraftChanged = (): void => {
    abortOutputRequests();
    generatedRevisionRef.current = null;
    controllerRef.current?.abort();
    controllerRef.current = null;
    dispatch({ type: 'draft-changed' });
  };

  const markTemplateDraft = (template: typeof fixture.template): void => {
    setDraftTemplate(template);
    markDraftChanged();
  };

  const markDataDraft = (event: ChangeEvent<HTMLTextAreaElement>): void => {
    setDraftData(event.currentTarget.value);
    setDataError('');
    markDraftChanged();
  };

  const applyDraft = (): void => {
    let data: unknown;
    try {
      data = JSON.parse(draftData);
    } catch {
      setDataError('Data must be valid JSON before regeneration.');
      return;
    }
    setDataError('');
    revisionRef.current += 1;
    startRender(
      revisionRef.current,
      draftTemplate,
      data,
      outputDocumentMetadata(draftTemplate, data),
    );
  };

  const pageCount = state.generatedDocument?.print.pages.length ?? 0;
  const canOutput =
    state.phase === 'ready' && state.generatedDocument !== null && state.generatedMetadata !== null;
  const runOutput = <Result,>(
    kind: OutputKind,
    action: (
      adapters: OutputStudioAdapters,
      generated: GeneratedDocument,
      signal: AbortSignal,
    ) => Promise<Result>,
    finish: (result: Result, metadata: OutputDocumentMetadata) => string,
  ): void => {
    const generated = state.generatedDocument;
    const metadata = state.generatedMetadata;
    const adapters = adaptersRef.current;
    const revision = state.generatedRevision;
    if (
      !canOutput ||
      generated === null ||
      metadata === null ||
      adapters === null ||
      revision === null
    ) {
      return;
    }
    outputRequestsRef.current.get(kind)?.controller.abort();
    const request: ActiveOutputRequest = {
      requestId: (outputRequestIdRef.current += 1),
      revision,
      controller: new AbortController(),
    };
    outputRequestsRef.current.set(kind, request);
    dispatch({ type: 'output-started', kind, requestId: request.requestId });
    const isCurrent = (): boolean =>
      mountedRef.current &&
      !request.controller.signal.aborted &&
      outputRequestsRef.current.get(kind) === request &&
      generatedRevisionRef.current === request.revision;
    void action(adapters, generated, request.controller.signal)
      .then((result) => {
        if (!isCurrent()) return;
        const message = finish(result, metadata);
        if (!isCurrent()) return;
        dispatch({
          type: 'output-finished',
          kind,
          requestId: request.requestId,
          message,
        });
      })
      .catch((error: unknown) => {
        if (!isCurrent()) return;
        dispatch({
          type: 'output-failed',
          kind,
          requestId: request.requestId,
          message: outputErrorMessage(kind, error),
        });
      })
      .finally(() => {
        if (outputRequestsRef.current.get(kind) === request) {
          outputRequestsRef.current.delete(kind);
        }
      });
  };
  const print = (): void => {
    runOutput(
      'print',
      async (adapters, generated, signal) => adapters.print.print(generated, { signal }),
      () => 'Print dialog opened',
    );
  };
  const downloadPdf = (): void => {
    runOutput(
      'pdf',
      async (adapters, generated, signal) =>
        adapters.pdf.render(generated, {
          pages: 'all',
          metadata: { title: state.generatedMetadata?.title },
          tagged: false,
          signal,
        }),
      (pdf, metadata) => {
        downloadBlob(pdf, outputFilename('pdf', metadata.invoiceId));
        return 'PDF downloaded';
      },
    );
  };
  const downloadPng = (): void => {
    const selectedPage = 0;
    runOutput(
      'png',
      async (adapters, generated, signal) =>
        adapters.image.render(generated, {
          format: 'png',
          pages: [selectedPage],
          background: '#ffffff',
          dpi: 144,
          signal,
        }),
      ([png], metadata) => {
        if (png === undefined) throw new Error('PNG adapter returned no page');
        downloadBlob(png, outputFilename('png', metadata.invoiceId, selectedPage));
        return 'PNG page 1 downloaded';
      },
    );
  };
  const downloadXlsx = (): void => {
    runOutput(
      'xlsx',
      async (adapters, generated, signal) =>
        adapters.xlsx.render(generated, {
          formulaMode: 'formula-and-cached-value',
          compatibility: 'excel',
          signal,
        }),
      (xlsx, metadata) => {
        downloadBlob(xlsx, outputFilename('xlsx', metadata.invoiceId));
        return 'XLSX downloaded';
      },
    );
  };
  const status =
    state.phase === 'dirty'
      ? 'Preview is stale. Apply & regenerate to update every output.'
      : state.phase === 'rendering'
        ? 'Generating the shared document…'
        : state.phase === 'blocked'
          ? 'Generation is blocked. Review the diagnostics.'
          : 'Preview and outputs use the current generated document.';

  return (
    <main className={styles.outputStudio}>
      <header className={styles.outputStudioHeader}>
        <div>
          <p className={styles.eyebrow}>Output playground</p>
          <h1>Output Studio</h1>
          <p>One document · many outputs.</p>
          <p>Preview, print, PDF, and PNG share one exact display list.</p>
          <p>XLSX uses the semantic workbook in the same GeneratedDocument.</p>
        </div>
        <p className={styles.revision}>
          {state.generatedRevision === null
            ? 'Preparing GeneratedDocument'
            : `GeneratedDocument · revision ${state.generatedRevision}`}
        </p>
      </header>

      <div className={styles.outputStudioGrid}>
        <section className={styles.outputInputs} aria-labelledby="output-inputs-heading">
          <h2 id="output-inputs-heading">Output inputs</h2>
          <p>Invoice data and template changes are prepared here before regeneration.</p>
          <button
            type="button"
            aria-expanded={workbenchOpen}
            aria-controls="template-workbench"
            onClick={() => setWorkbenchOpen((open) => !open)}
          >
            Edit template
          </button>
          {workbenchOpen ? (
            <section
              id="template-workbench"
              className={styles.templateWorkbench}
              aria-labelledby="template-workbench-heading"
            >
              <h3 id="template-workbench-heading">Template workbench</h3>
              <p>
                Changes remain drafts until you explicitly apply them to a new generated revision.
              </p>
              <label className={styles.dataField}>
                Data JSON
                <textarea value={draftData} onChange={markDataDraft} rows={12} spellCheck={false} />
              </label>
              {dataError === '' ? null : <p role="alert">{dataError}</p>}
              <div className={styles.templateSheet} data-mode="template">
                <TegoSheet
                  document={fixture.document}
                  mode="template"
                  template={draftTemplate}
                  onTemplateChange={markTemplateDraft}
                />
              </div>
              <button type="button" onClick={applyDraft} disabled={state.phase !== 'dirty'}>
                Apply &amp; regenerate
              </button>
            </section>
          ) : null}
        </section>

        <section className={styles.exactPreview} aria-labelledby="exact-preview-heading">
          <h2 id="exact-preview-heading">Exact page preview</h2>
          {state.generatedDocument === null ? (
            <p role="status">Preparing deterministic invoice pages…</p>
          ) : (
            <TemplatePreview document={state.generatedDocument} />
          )}
        </section>

        <section className={styles.pipelineOutputs} aria-labelledby="pipeline-outputs-heading">
          <h2 id="pipeline-outputs-heading">Pipeline and outputs</h2>
          <p>Every output starts from the same immutable GeneratedDocument.</p>
          <p role="status" aria-live="polite">
            {status}
          </p>
          {state.diagnostics.length === 0 ? null : (
            <ul className={styles.outputDiagnostics} aria-label="Generation diagnostics">
              {state.diagnostics.map((diagnostic, index) => (
                <li key={`${diagnostic.code}-${index}`}>{diagnostic.message}</li>
              ))}
            </ul>
          )}
          <div className={styles.outputActions}>
            <div className={styles.outputAction}>
              <button
                type="button"
                onClick={print}
                disabled={!canOutput || state.outputs.print.status === 'busy'}
                aria-describedby={
                  state.outputs.print.message === '' ? undefined : 'output-result-print'
                }
              >
                Print {pageCount} pages
              </button>
              {state.outputs.print.message === '' ? null : (
                <p
                  id="output-result-print"
                  role={state.outputs.print.status === 'error' ? 'alert' : 'status'}
                  aria-live={state.outputs.print.status === 'error' ? 'assertive' : 'polite'}
                >
                  {state.outputs.print.message}
                </p>
              )}
            </div>
            <div className={styles.outputAction}>
              <button
                type="button"
                onClick={downloadPdf}
                disabled={!canOutput || state.outputs.pdf.status === 'busy'}
                aria-describedby={
                  state.outputs.pdf.message === '' ? undefined : 'output-result-pdf'
                }
              >
                Download PDF
              </button>
              {state.outputs.pdf.message === '' ? null : (
                <p
                  id="output-result-pdf"
                  role={state.outputs.pdf.status === 'error' ? 'alert' : 'status'}
                  aria-live={state.outputs.pdf.status === 'error' ? 'assertive' : 'polite'}
                >
                  {state.outputs.pdf.message}
                </p>
              )}
            </div>
            <div className={styles.outputAction}>
              <button
                type="button"
                onClick={downloadPng}
                disabled={!canOutput || state.outputs.png.status === 'busy'}
                aria-describedby={
                  state.outputs.png.message === '' ? undefined : 'output-result-png'
                }
              >
                Download PNG page 1
              </button>
              {state.outputs.png.message === '' ? null : (
                <p
                  id="output-result-png"
                  role={state.outputs.png.status === 'error' ? 'alert' : 'status'}
                  aria-live={state.outputs.png.status === 'error' ? 'assertive' : 'polite'}
                >
                  {state.outputs.png.message}
                </p>
              )}
            </div>
            <div className={styles.outputAction}>
              <button
                type="button"
                onClick={downloadXlsx}
                disabled={!canOutput || state.outputs.xlsx.status === 'busy'}
                aria-describedby={
                  state.outputs.xlsx.message === '' ? undefined : 'output-result-xlsx'
                }
              >
                Download XLSX
              </button>
              {state.outputs.xlsx.message === '' ? null : (
                <p
                  id="output-result-xlsx"
                  role={state.outputs.xlsx.status === 'error' ? 'alert' : 'status'}
                  aria-live={state.outputs.xlsx.status === 'error' ? 'assertive' : 'polite'}
                >
                  {state.outputs.xlsx.message}
                </p>
              )}
            </div>
          </div>
        </section>
      </div>
    </main>
  );
}
