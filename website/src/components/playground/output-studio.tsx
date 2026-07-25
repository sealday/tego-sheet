import { TegoSheet, TemplatePreview, type Diagnostic } from 'tego-sheet';
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
import { createInvoiceOutputFixture } from './output-studio-fixtures';
import { createOutputStudioState, reduceOutputStudioState } from './output-studio-model';
import { renderOutputRevision } from './output-studio-pipeline';
import styles from './playground.module.css';

export function OutputStudio(): ReactElement {
  const [fixture] = useState(createInvoiceOutputFixture);
  const [draftTemplate, setDraftTemplate] = useState(fixture.template);
  const [draftData, setDraftData] = useState(() => JSON.stringify(fixture.data, null, 2));
  const [workbenchOpen, setWorkbenchOpen] = useState(false);
  const [dataError, setDataError] = useState('');
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
    const controller = startRender(1, fixture.template, fixture.data);
    return () => controller.abort();
  }, [fixture, startRender]);

  const markTemplateDraft = (template: typeof fixture.template): void => {
    setDraftTemplate(template);
    dispatch({ type: 'draft-changed' });
  };

  const markDataDraft = (event: ChangeEvent<HTMLTextAreaElement>): void => {
    setDraftData(event.currentTarget.value);
    setDataError('');
    dispatch({ type: 'draft-changed' });
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
          <p>
            One document · many outputs. The exact generated display list shown here is the shared
            artifact for print, PDF, images, and spreadsheets.
          </p>
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
            <button type="button" disabled={state.phase !== 'ready'}>
              Print {pageCount} pages
            </button>
            <button type="button" disabled={state.phase !== 'ready'}>
              Export PDF
            </button>
            <button type="button" disabled={state.phase !== 'ready'}>
              Export PNG
            </button>
            <button type="button" disabled={state.phase !== 'ready'}>
              Export XLSX
            </button>
          </div>
        </section>
      </div>
    </main>
  );
}
