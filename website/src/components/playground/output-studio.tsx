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

export function OutputStudio({ adapters: injectedAdapters }: OutputStudioProps = {}): ReactElement {
  const [fixture] = useState(createInvoiceOutputFixture);
  const [adapters] = useState(() => injectedAdapters ?? createOutputStudioAdapters());
  const [draftTemplate, setDraftTemplate] = useState(fixture.template);
  const [draftData, setDraftData] = useState(() => JSON.stringify(fixture.data, null, 2));
  const [workbenchOpen, setWorkbenchOpen] = useState(false);
  const [dataError, setDataError] = useState('');
  const [busyOutputs, setBusyOutputs] = useState<ReadonlySet<OutputKind>>(() => new Set());
  const [state, dispatch] = useReducer(reduceOutputStudioState, undefined, createOutputStudioState);
  const revisionRef = useRef(1);
  const controllerRef = useRef<AbortController | null>(null);

  const startRender = useCallback(
    (revision: number, template: typeof fixture.template, data: unknown): AbortController => {
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
            dispatch({
              type: 'render-succeeded',
              revision,
              document: result.document,
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
    [fixture],
  );

  useEffect(() => {
    startRender(1, fixture.template, fixture.data);
    return () => controllerRef.current?.abort();
  }, [fixture, startRender]);

  useEffect(
    () => () => {
      if (injectedAdapters === undefined) adapters.print.dispose();
    },
    [adapters, injectedAdapters],
  );

  const markDraftChanged = (): void => {
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
    startRender(revisionRef.current, draftTemplate, data);
  };

  const pageCount = state.generatedDocument?.print.pages.length ?? 0;
  const canOutput = state.phase === 'ready' && state.generatedDocument !== null;
  const runOutput = async (
    kind: OutputKind,
    action: (generated: GeneratedDocument) => Promise<string>,
  ): Promise<void> => {
    const generated = state.generatedDocument;
    if (!canOutput || generated === null) return;
    setBusyOutputs((current) => new Set(current).add(kind));
    dispatch({ type: 'output-started', kind });
    try {
      dispatch({ type: 'output-finished', message: await action(generated) });
    } catch (error: unknown) {
      dispatch({ type: 'output-failed', message: outputErrorMessage(kind, error) });
    } finally {
      setBusyOutputs((current) => {
        const next = new Set(current);
        next.delete(kind);
        return next;
      });
    }
  };
  const print = (): void => {
    void runOutput('print', async (generated) => {
      await adapters.print.print(generated);
      return 'Print dialog opened';
    });
  };
  const downloadPdf = (): void => {
    void runOutput('pdf', async (generated) => {
      const pdf = await adapters.pdf.render(generated, {
        pages: 'all',
        metadata: { title: fixture.template.name },
        tagged: false,
      });
      downloadBlob(pdf, outputFilename('pdf', fixture.data.invoice.id));
      return 'PDF downloaded';
    });
  };
  const downloadPng = (): void => {
    const selectedPage = 0;
    void runOutput('png', async (generated) => {
      const [png] = await adapters.image.render(generated, {
        format: 'png',
        pages: [selectedPage],
        background: '#ffffff',
        dpi: 144,
      });
      if (png === undefined) throw new Error('PNG adapter returned no page');
      downloadBlob(png, outputFilename('png', fixture.data.invoice.id, selectedPage));
      return 'PNG page 1 downloaded';
    });
  };
  const downloadXlsx = (): void => {
    void runOutput('xlsx', async (generated) => {
      const xlsx = await adapters.xlsx.render(generated, {
        formulaMode: 'formula-and-cached-value',
        compatibility: 'excel',
      });
      downloadBlob(xlsx, outputFilename('xlsx', fixture.data.invoice.id));
      return 'XLSX downloaded';
    });
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
          {state.outputMessage === '' ? null : (
            <p
              role={state.outputMessage.includes(' failed') ? 'alert' : 'status'}
              aria-live={state.outputMessage.includes(' failed') ? 'assertive' : 'polite'}
            >
              {state.outputMessage}
            </p>
          )}
          {state.diagnostics.length === 0 ? null : (
            <ul className={styles.outputDiagnostics} aria-label="Generation diagnostics">
              {state.diagnostics.map((diagnostic, index) => (
                <li key={`${diagnostic.code}-${index}`}>{diagnostic.message}</li>
              ))}
            </ul>
          )}
          <div className={styles.outputActions}>
            <button type="button" onClick={print} disabled={!canOutput || busyOutputs.has('print')}>
              Print {pageCount} pages
            </button>
            <button
              type="button"
              onClick={downloadPdf}
              disabled={!canOutput || busyOutputs.has('pdf')}
            >
              Download PDF
            </button>
            <button
              type="button"
              onClick={downloadPng}
              disabled={!canOutput || busyOutputs.has('png')}
            >
              Download PNG page 1
            </button>
            <button
              type="button"
              onClick={downloadXlsx}
              disabled={!canOutput || busyOutputs.has('xlsx')}
            >
              Download XLSX
            </button>
          </div>
        </section>
      </div>
    </main>
  );
}
