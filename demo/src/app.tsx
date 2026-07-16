import {
  Component,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import {
  TegoSheet,
  type ActiveSheetChangeEvent,
  type Selection,
  type TegoSheetError,
  type TegoSheetHandle,
  type WorkbookChange,
  type WorkbookData,
  type WorkbookInput,
} from 'tego-sheet';
import { zhCN } from 'tego-sheet/locales/zh-cn';
import {
  appendPreviewEvent,
  cloneExampleWorkbook,
  formatWorkbookJson,
  parseWorkbookJson,
  type PreviewEventInput,
} from './workbench-model';

type PreviewMode = 'uncontrolled' | 'controlled';
type LocaleCode = 'en' | 'zh-CN';

interface PreviewErrorBoundaryProps {
  readonly children: ReactNode;
  readonly onError: (error: Error) => void;
}

interface PreviewErrorBoundaryState {
  readonly failed: boolean;
}

class PreviewErrorBoundary extends Component<PreviewErrorBoundaryProps, PreviewErrorBoundaryState> {
  state: PreviewErrorBoundaryState = { failed: false };

  static getDerivedStateFromError(): PreviewErrorBoundaryState {
    return { failed: true };
  }

  componentDidCatch(error: Error): void {
    this.props.onError(error);
  }

  render(): ReactNode {
    return this.state.failed ? null : this.props.children;
  }
}

function eventDetails(value: unknown): string {
  return JSON.stringify(value);
}

function workbookStatus(workbook: WorkbookInput): string {
  const sheets = Array.isArray(workbook) ? workbook : [workbook];
  const firstName = sheets[0]?.name?.trim() || 'Untitled';
  const suffix = sheets.length === 1 ? 'sheet' : 'sheets';
  return `${firstName} · ${sheets.length} ${suffix}`;
}

export function App() {
  const initialWorkbook = useMemo(cloneExampleWorkbook, []);
  const sheetRef = useRef<TegoSheetHandle>(null);
  const lastStableRef = useRef<{ workbook: WorkbookData; mode: PreviewMode }>({
    workbook: initialWorkbook,
    mode: 'uncontrolled',
  });
  const recoveringRef = useRef(false);
  const [mode, setMode] = useState<PreviewMode>('uncontrolled');
  const [readOnly, setReadOnly] = useState(false);
  const [localeCode, setLocaleCode] = useState<LocaleCode>('en');
  const [workbook, setWorkbook] = useState<WorkbookInput>(initialWorkbook);
  const [jsonText, setJsonText] = useState(() => formatWorkbookJson(initialWorkbook));
  const [mountEpoch, setMountEpoch] = useState(0);
  const [controlsCollapsed, setControlsCollapsed] = useState(false);
  const [jsonVisible, setJsonVisible] = useState(false);
  const [eventsVisible, setEventsVisible] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [events, setEvents] = useState<ReturnType<typeof appendPreviewEvent>>([]);

  const locale = useMemo(() => (localeCode === 'zh-CN' ? zhCN : undefined), [localeCode]);

  useEffect(() => {
    const stableWorkbook = sheetRef.current?.getValue();
    if (stableWorkbook === undefined) return;

    lastStableRef.current = { workbook: stableWorkbook, mode };
    recoveringRef.current = false;
  }, [mode, mountEpoch, workbook]);

  const recordEvent = useCallback((input: Omit<PreviewEventInput, 'timestamp'>) => {
    setEvents((current) =>
      appendPreviewEvent(current, {
        ...input,
        timestamp: new Date().toISOString(),
      }),
    );
  }, []);

  const changeMode = (nextMode: PreviewMode) => {
    recoveringRef.current = false;
    setMode(nextMode);
    setMountEpoch((current) => current + 1);
  };

  const importJson = () => {
    try {
      const imported = parseWorkbookJson(jsonText);
      const currentWorkbook = sheetRef.current?.getValue();
      if (currentWorkbook !== undefined) {
        lastStableRef.current = { workbook: currentWorkbook, mode };
      }
      recoveringRef.current = false;
      setWorkbook(imported);
      setJsonText(formatWorkbookJson(imported));
      setError(null);
      setMountEpoch((current) => current + 1);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Workbook JSON is invalid.');
    }
  };

  const resetWorkbook = () => {
    const freshWorkbook = cloneExampleWorkbook();
    recoveringRef.current = false;
    setWorkbook(freshWorkbook);
    setJsonText(formatWorkbookJson(freshWorkbook));
    setError(null);
    setEvents([]);
    setMountEpoch((current) => current + 1);
  };

  const exportJson = () => {
    const exported = sheetRef.current?.getValue();
    if (exported === undefined) {
      setError('The spreadsheet is not ready to export.');
      return;
    }

    setJsonText(formatWorkbookJson(exported));
    setError(null);
  };

  const handleChange = (nextWorkbook: WorkbookData, change: WorkbookChange) => {
    setWorkbook(nextWorkbook);
    recordEvent({ label: 'Workbook changed', details: eventDetails(change) });
  };

  const handleActiveSheetChange = (change: ActiveSheetChangeEvent) => {
    recordEvent({ label: 'Active sheet changed', details: eventDetails(change) });
  };

  const handleSelectionChange = (selection: Selection) => {
    recordEvent({ label: 'Selection changed', details: eventDetails(selection) });
  };

  const handleError = (sheetError: TegoSheetError) => {
    setError(sheetError.message);
    recordEvent({ label: 'Spreadsheet error', details: sheetError.message });
  };

  const recoverFromRenderError = (renderError: Error) => {
    if (recoveringRef.current) {
      setError(`Spreadsheet recovery failed: ${renderError.message}`);
      return;
    }

    recoveringRef.current = true;
    const stable = lastStableRef.current;
    setWorkbook(stable.workbook);
    setMode(stable.mode);
    setJsonText(formatWorkbookJson(stable.workbook));
    setError(`Workbook import failed: ${renderError.message}`);
    setMountEpoch((current) => current + 1);
  };

  return (
    <main className="preview-shell">
      <header className="preview-controls" aria-label="Preview controls">
        <div className="preview-controls__summary">
          <div>
            <h1>tego-sheet workbench</h1>
            <p>
              <span>Mode: {mode === 'controlled' ? 'Controlled' : 'Uncontrolled'}</span>
              {' · '}
              <span>Workbook: {workbookStatus(workbook)}</span>
            </p>
          </div>
          <button
            type="button"
            aria-expanded={!controlsCollapsed}
            aria-controls="preview-controls-secondary"
            onClick={() => setControlsCollapsed((current) => !current)}
          >
            {controlsCollapsed ? 'Expand controls' : 'Collapse controls'}
          </button>
        </div>

        {error !== null && (
          <p className="preview-alert" role="alert">
            {error}
          </p>
        )}

        {!controlsCollapsed && (
          <div id="preview-controls-secondary" className="preview-controls__secondary">
            <div className="preview-controls__fields">
              <label>
                Mode
                <select
                  value={mode}
                  onChange={(event) => changeMode(event.currentTarget.value as PreviewMode)}
                >
                  <option value="uncontrolled">Uncontrolled</option>
                  <option value="controlled">Controlled</option>
                </select>
              </label>

              <label>
                <input
                  type="checkbox"
                  checked={readOnly}
                  onChange={(event) => setReadOnly(event.currentTarget.checked)}
                />
                Read only
              </label>

              <label>
                Locale
                <select
                  value={localeCode}
                  onChange={(event) => setLocaleCode(event.currentTarget.value as LocaleCode)}
                >
                  <option value="en">English</option>
                  <option value="zh-CN">简体中文</option>
                </select>
              </label>
            </div>

            <div className="preview-controls__actions">
              <button type="button" onClick={resetWorkbook}>
                Reset workbook
              </button>
              <button type="button" onClick={importJson}>
                Import JSON
              </button>
              <button type="button" onClick={exportJson}>
                Export JSON
              </button>
              <button
                type="button"
                aria-expanded={jsonVisible}
                aria-controls="workbook-json-panel"
                onClick={() => setJsonVisible((current) => !current)}
              >
                {jsonVisible ? 'Hide JSON' : 'Show JSON'}
              </button>
              <button
                type="button"
                aria-expanded={eventsVisible}
                aria-controls="preview-events-panel"
                onClick={() => setEventsVisible((current) => !current)}
              >
                {eventsVisible ? 'Hide events' : 'Show events'}
              </button>
            </div>

            {jsonVisible && (
              <div id="workbook-json-panel" className="preview-json-panel">
                <label htmlFor="workbook-json">Workbook JSON</label>
                <textarea
                  id="workbook-json"
                  value={jsonText}
                  onChange={(event) => setJsonText(event.currentTarget.value)}
                />
              </div>
            )}

            {eventsVisible && (
              <ol
                id="preview-events-panel"
                className="preview-events-panel"
                role="log"
                aria-label="Spreadsheet events"
              >
                {events.map((event) => (
                  <li key={event.id}>
                    <strong>{event.label}</strong>
                    {event.details === undefined ? null : ` — ${event.details}`}
                    <time dateTime={event.timestamp}>{event.timestamp}</time>
                  </li>
                ))}
              </ol>
            )}
          </div>
        )}
      </header>

      <section className="preview-workspace" aria-label="Spreadsheet preview">
        <PreviewErrorBoundary
          key={`preview-boundary-${mountEpoch}`}
          onError={recoverFromRenderError}
        >
          <TegoSheet
            key={mountEpoch}
            ref={sheetRef}
            {...(mode === 'controlled' ? { value: workbook } : { defaultValue: workbook })}
            readOnly={readOnly}
            locale={locale}
            onChange={handleChange}
            onActiveSheetChange={handleActiveSheetChange}
            onSelectionChange={handleSelectionChange}
            onError={handleError}
          />
        </PreviewErrorBoundary>
      </section>
    </main>
  );
}
