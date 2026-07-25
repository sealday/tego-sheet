import { TegoSheet, type SpreadsheetDocument, type TegoSheetHandle } from 'tego-sheet';
import { en } from 'tego-sheet/locales/en';
import { zhCN } from 'tego-sheet/locales/zh-cn';
import { de } from 'tego-sheet/locales/de';
import { nl } from 'tego-sheet/locales/nl';
import 'tego-sheet/styles.css';
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ComponentProps,
  type KeyboardEvent,
  type ReactElement,
} from 'react';
import { PLAYGROUND_LOCALES, PLAYGROUND_PRESETS, createFixture } from './playground-fixtures';
import {
  appendPlaygroundEvent,
  type PlaygroundCallbackName,
  type PlaygroundEvent,
  type PlaygroundMode,
} from './playground-model';
import { PlaygroundErrorBoundary } from './playground-error-boundary';
import {
  readPlaygroundLocation,
  writePlaygroundLocation,
  type PlaygroundLocation,
  type PlaygroundWorkspace,
} from './playground-workspace';
import { OutputStudio } from './output-studio';
import styles from './playground.module.css';

type SheetProps = ComponentProps<typeof TegoSheet>;
type SheetCallbacks = Pick<
  SheetProps,
  | 'onDocumentChange'
  | 'onActiveSheetChange'
  | 'onSelectionChange'
  | 'onCellEdit'
  | 'onPaste'
  | 'onError'
>;
type ToolbarRenderer = Exclude<SheetProps['toolbar'], 'default' | false | undefined>;
type SheetTabsRenderer = Exclude<SheetProps['sheetTabs'], 'default' | false | undefined>;

const LOCALES = { en, 'zh-CN': zhCN, de, nl } as const;
type LocaleId = keyof typeof LOCALES;
const WORKSPACE_OPTIONS = Object.freeze([
  { id: 'spreadsheet', label: 'Spreadsheet' },
  { id: 'output', label: 'Output Studio' },
] as const);

function toPublicJson(value: unknown, seen = new WeakSet<object>()): PlaygroundEvent['payload'] {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number') return Number.isFinite(value) ? value : String(value);
  if (typeof value !== 'object') return String(value);
  if (seen.has(value)) return '[Circular]';

  seen.add(value);
  const json = Array.isArray(value)
    ? value.map((entry) => toPublicJson(entry, seen))
    : Object.fromEntries(
        Object.entries(value).map(([key, entry]) => [key, toPublicJson(entry, seen)]),
      );
  seen.delete(value);
  return json;
}

function sitePath(path: string): string {
  return `${window.location.pathname.replace(/\/playground\/?$/, '')}${path}`;
}

function reloadWindow(): void {
  window.location.reload();
}

export interface PlaygroundProps {
  readonly onReload?: () => void;
}

interface SpreadsheetWorkspaceProps extends PlaygroundProps {
  readonly historyRevision: number;
  readonly mode: PlaygroundMode;
  readonly onModeChange: (mode: PlaygroundMode) => void;
  readonly onRecover: () => void;
  readonly restoredFromHistory: boolean;
}

interface PresetSheetProps {
  readonly fixture: SpreadsheetDocument;
  readonly locale: (typeof LOCALES)[LocaleId];
  readonly callbacks: SheetCallbacks;
  readonly sheetRef: React.RefObject<TegoSheetHandle | null>;
}

function ControlledSheet({ fixture, locale, callbacks, sheetRef }: PresetSheetProps): ReactElement {
  const [value, setValue] = useState<SpreadsheetDocument>(fixture);
  const onChange: NonNullable<SheetCallbacks['onDocumentChange']> = (nextValue, change) => {
    setValue(nextValue);
    callbacks.onDocumentChange?.(nextValue, change);
  };

  return (
    <TegoSheet
      ref={sheetRef}
      document={value}
      locale={locale}
      {...callbacks}
      onDocumentChange={onChange}
    />
  );
}

function CustomToolbar(props: Parameters<ToolbarRenderer>[0]) {
  return (
    <div className={styles.customChrome} aria-label="Custom toolbar">
      <strong>Roadmap tools</strong>
      <button
        type="button"
        disabled={props.disabledActions.has('set-style')}
        onClick={() => props.execute({ type: 'set-style', patch: { font: { bold: true } } })}
      >
        Bold selection
      </button>
    </div>
  );
}

function CustomSheetTabs(props: Parameters<SheetTabsRenderer>[0]) {
  return (
    <div className={styles.customChrome} aria-label="Custom sheet tabs">
      {props.sheets.map((sheet) => (
        <button
          key={sheet.id}
          type="button"
          aria-pressed={sheet.id === props.activeSheet}
          onClick={() => props.activate(sheet.id)}
        >
          {sheet.name}
        </button>
      ))}
      <button type="button" disabled={props.readOnly} onClick={() => props.add('Demo sheet')}>
        Add demo sheet
      </button>
    </div>
  );
}

function UncontrolledSheet({
  mode,
  fixture,
  locale,
  callbacks,
  sheetRef,
}: PresetSheetProps & { readonly mode: PlaygroundMode }): ReactElement {
  return (
    <TegoSheet
      ref={sheetRef}
      defaultDocument={fixture}
      locale={locale}
      toolbar={mode === 'custom-chrome' ? CustomToolbar : 'default'}
      sheetTabs={mode === 'custom-chrome' ? CustomSheetTabs : 'default'}
      {...callbacks}
    />
  );
}

interface PresetSessionProps {
  readonly mode: PlaygroundMode;
  readonly presetKey: string;
  readonly setStatus: (message: string) => void;
  readonly onReset: () => void;
}

function PresetSession({ mode, presetKey, setStatus, onReset }: PresetSessionProps): ReactElement {
  const preset = PLAYGROUND_PRESETS[mode];
  const fixture = useMemo(() => createFixture(mode), [mode]);
  const sheetRef = useRef<TegoSheetHandle>(null);
  const sequence = useRef(0);
  const copyRequest = useRef(0);
  const mounted = useRef(false);
  const [events, setEvents] = useState<readonly PlaygroundEvent[]>([]);
  const [snapshot, setSnapshot] = useState<SpreadsheetDocument>(fixture);
  const [localeId, setLocaleId] = useState<LocaleId>('en');

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      copyRequest.current += 1;
    };
  }, []);

  const record = useCallback((callback: PlaygroundCallbackName, payload: unknown): void => {
    const nextSequence = ++sequence.current;
    setEvents((current) =>
      appendPlaygroundEvent(current, {
        sequence: nextSequence,
        callback,
        payload: toPublicJson(payload),
      }),
    );
  }, []);

  const callbacks = useMemo<SheetCallbacks>(
    () => ({
      onDocumentChange: (value, change) => {
        setSnapshot(value);
        record('onDocumentChange', change);
      },
      onActiveSheetChange: (event) => record('onActiveSheetChange', event),
      onSelectionChange: (event) => record('onSelectionChange', event),
      onCellEdit: (event) => record('onCellEdit', event),
      onPaste: (event) => record('onPaste', event),
      onError: (error) =>
        record('onError', {
          code: error.code,
          message: error.message,
          recoverable: error.recoverable,
        }),
    }),
    [record],
  );

  const formattedSnapshot = useMemo(() => JSON.stringify(snapshot, null, 2), [snapshot]);
  const announceNewStatus = useCallback(
    (message: string): void => {
      copyRequest.current += 1;
      setStatus(message);
    },
    [setStatus],
  );
  const copySnapshot = async (): Promise<void> => {
    const request = ++copyRequest.current;
    try {
      await navigator.clipboard.writeText(formattedSnapshot);
      if (mounted.current && request === copyRequest.current) {
        setStatus('Document JSON copied');
      }
    } catch {
      if (mounted.current && request === copyRequest.current) {
        setStatus('Could not copy workbook JSON');
      }
    }
  };
  const refreshSnapshot = (): void => {
    const currentValue = sheetRef.current?.getDocument();
    if (currentValue) setSnapshot(currentValue);
    announceNewStatus('Document JSON refreshed from TegoSheetHandle.getDocument()');
  };

  return (
    <section
      className={styles.presetBoundary}
      data-preset-key={presetKey}
      data-testid="preset-boundary"
    >
      <div className={styles.sheetPanel}>
        {mode === 'locales' ? (
          <label className={styles.field}>
            <span>Locale</span>
            <select
              value={localeId}
              onChange={(event) => setLocaleId(event.currentTarget.value as LocaleId)}
            >
              {PLAYGROUND_LOCALES.map((locale) => {
                const id = locale.subpath.endsWith('/zh-cn')
                  ? 'zh-CN'
                  : locale.subpath.slice(locale.subpath.lastIndexOf('/') + 1);
                return (
                  <option key={locale.subpath} value={id}>
                    {locale.label}
                  </option>
                );
              })}
            </select>
          </label>
        ) : null}
        <div className={styles.sheetHost}>
          {mode === 'controlled' ? (
            <ControlledSheet
              fixture={fixture}
              locale={LOCALES[localeId]}
              callbacks={callbacks}
              sheetRef={sheetRef}
            />
          ) : (
            <UncontrolledSheet
              mode={mode}
              fixture={fixture}
              locale={LOCALES[localeId]}
              callbacks={callbacks}
              sheetRef={sheetRef}
            />
          )}
        </div>
      </div>

      <aside className={styles.inspector} aria-label="Playground inspector">
        <div>
          <p className={styles.eyebrow}>Current preset</p>
          <h2>{preset.label}</h2>
          <p>{preset.description}</p>
          <a href={sitePath(preset.docsLink)}>Read the guide</a>
        </div>
        <div>
          <h3>Public APIs</h3>
          <ul className={styles.apiList}>
            {preset.publicApis.map((api) => (
              <li key={api}>
                <code>{api}</code>
              </li>
            ))}
          </ul>
        </div>
        <button type="button" onClick={onReset}>
          Reset mode
        </button>
        <div>
          <div className={styles.inspectorHeading}>
            <h3>Document JSON</h3>
            <div className={styles.buttonRow}>
              <button type="button" onClick={refreshSnapshot}>
                Refresh JSON
              </button>
              <button type="button" onClick={() => void copySnapshot()}>
                Copy JSON
              </button>
            </div>
          </div>
          <pre className={styles.json} aria-label="Document JSON" tabIndex={0}>
            {formattedSnapshot}
          </pre>
        </div>
        <div>
          <h3>Events</h3>
          {events.length === 0 ? (
            <p className={styles.emptyEvents}>Interact with the sheet to inspect callbacks.</p>
          ) : (
            <ol className={styles.events}>
              {[...events].reverse().map((event) => (
                <li key={event.sequence} aria-label={`Event ${event.sequence}`}>
                  <strong>
                    #{event.sequence} {event.callback}
                  </strong>
                  <pre>{JSON.stringify(event.payload, null, 2)}</pre>
                </li>
              ))}
            </ol>
          )}
        </div>
      </aside>
    </section>
  );
}

function SpreadsheetWorkspace({
  historyRevision,
  mode,
  onModeChange,
  onRecover,
  onReload = reloadWindow,
  restoredFromHistory,
}: SpreadsheetWorkspaceProps): ReactElement {
  const [resetRevision, setResetRevision] = useState(0);
  const [status, setStatusState] = useState({
    historyRevision: restoredFromHistory ? historyRevision - 1 : historyRevision,
    message: '',
  });
  const setStatus = useCallback(
    (message: string): void => setStatusState({ historyRevision, message }),
    [historyRevision],
  );
  const displayedStatus =
    restoredFromHistory && status.historyRevision !== historyRevision
      ? `${PLAYGROUND_PRESETS[mode].label} restored from browser history`
      : status.message;

  const selectMode = (nextMode: PlaygroundMode): void => {
    if (nextMode === mode) return;
    onModeChange(nextMode);
    setResetRevision(0);
    setStatus(`${PLAYGROUND_PRESETS[nextMode].label} selected`);
  };

  const resetMode = (): void => {
    setResetRevision((revision) => revision + 1);
    setStatus(`${PLAYGROUND_PRESETS[mode].label} reset`);
  };

  const recoverFromError = (): void => {
    onRecover();
    setResetRevision((revision) => revision + 1);
    setStatus('Playground reset to Uncontrolled');
  };

  const presetKey = `${mode}:${resetRevision}:${historyRevision}`;

  return (
    <>
      <header className={styles.header}>
        <div>
          <p>Switch presets without leaving the page, then inspect document data and callbacks.</p>
        </div>
        <fieldset className={styles.modePicker}>
          <legend>Playground mode</legend>
          {Object.values(PLAYGROUND_PRESETS).map((preset) => (
            <label key={preset.mode}>
              <input
                type="radio"
                name="playground-mode"
                value={preset.mode}
                checked={mode === preset.mode}
                onChange={() => selectMode(preset.mode)}
              />
              <span>{preset.label}</span>
            </label>
          ))}
        </fieldset>
      </header>
      <p className={styles.srStatus} role="status" aria-live="polite">
        {displayedStatus}
      </p>
      <PlaygroundErrorBoundary key={presetKey} onReset={recoverFromError} onReload={onReload}>
        <PresetSession
          key={presetKey}
          mode={mode}
          presetKey={presetKey}
          setStatus={setStatus}
          onReset={resetMode}
        />
      </PlaygroundErrorBoundary>
    </>
  );
}

export function Playground({ onReload = reloadWindow }: PlaygroundProps = {}): ReactElement {
  const [location, setLocation] = useState<PlaygroundLocation>(() =>
    readPlaygroundLocation(window.location.search),
  );
  const [historyRevision, setHistoryRevision] = useState(0);
  const [restoredFromHistory, setRestoredFromHistory] = useState(false);

  const urlForLocation = useCallback(
    (nextLocation: PlaygroundLocation): string =>
      `${writePlaygroundLocation(
        window.location.pathname,
        window.location.search,
        nextLocation,
      )}${window.location.hash}`,
    [],
  );

  const replaceLocation = useCallback(
    (nextLocation: PlaygroundLocation): void => {
      window.history.replaceState(window.history.state, '', urlForLocation(nextLocation));
      setLocation(nextLocation);
      setRestoredFromHistory(false);
    },
    [urlForLocation],
  );

  useEffect(() => {
    const canonicalUrl = urlForLocation(location);
    if (
      `${window.location.pathname}${window.location.search}${window.location.hash}` !== canonicalUrl
    ) {
      window.history.replaceState(window.history.state, '', canonicalUrl);
    }
  }, [location, urlForLocation]);

  useEffect(() => {
    const onPopState = (): void => {
      const nextLocation = readPlaygroundLocation(window.location.search);
      const canonicalUrl = urlForLocation(nextLocation);
      if (
        `${window.location.pathname}${window.location.search}${window.location.hash}` !==
        canonicalUrl
      ) {
        window.history.replaceState(window.history.state, '', canonicalUrl);
      }
      setLocation(nextLocation);
      setHistoryRevision((revision) => revision + 1);
      setRestoredFromHistory(true);
    };
    window.addEventListener('popstate', onPopState);
    return () => window.removeEventListener('popstate', onPopState);
  }, [urlForLocation]);

  const pushLocation = (nextLocation: PlaygroundLocation): void => {
    window.history.pushState(window.history.state, '', urlForLocation(nextLocation));
    setLocation(nextLocation);
    setRestoredFromHistory(false);
  };

  const selectWorkspace = (workspace: PlaygroundWorkspace): void => {
    if (workspace === location.workspace) return;
    pushLocation({ ...location, workspace });
  };

  const navigateWorkspaceTabs = (
    event: KeyboardEvent<HTMLButtonElement>,
    currentIndex: number,
  ): void => {
    let nextIndex: number;
    if (event.key === 'ArrowRight') nextIndex = (currentIndex + 1) % WORKSPACE_OPTIONS.length;
    else if (event.key === 'ArrowLeft')
      nextIndex = (currentIndex - 1 + WORKSPACE_OPTIONS.length) % WORKSPACE_OPTIONS.length;
    else if (event.key === 'Home') nextIndex = 0;
    else if (event.key === 'End') nextIndex = WORKSPACE_OPTIONS.length - 1;
    else return;

    event.preventDefault();
    const next = WORKSPACE_OPTIONS[nextIndex]!;
    const tabs =
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]');
    tabs?.[nextIndex]?.focus();
    selectWorkspace(next.id);
  };

  return (
    <main className={styles.playground}>
      <header className={styles.workspaceHeader}>
        <div>
          <p className={styles.eyebrow}>Live public API examples</p>
          <h1>Playground</h1>
        </div>
        <div className={styles.workspaceTabs} role="tablist" aria-label="Playground workspace">
          {WORKSPACE_OPTIONS.map((workspace, index) => (
            <button
              key={workspace.id}
              id={`workspace-tab-${workspace.id}`}
              role="tab"
              type="button"
              aria-controls={`workspace-panel-${workspace.id}`}
              aria-selected={location.workspace === workspace.id}
              tabIndex={location.workspace === workspace.id ? 0 : -1}
              onClick={() => selectWorkspace(workspace.id)}
              onKeyDown={(event) => navigateWorkspaceTabs(event, index)}
            >
              {workspace.label}
            </button>
          ))}
        </div>
      </header>
      <section
        id="workspace-panel-spreadsheet"
        className={styles.workspacePanel}
        role="tabpanel"
        aria-labelledby="workspace-tab-spreadsheet"
        hidden={location.workspace !== 'spreadsheet'}
      >
        {location.workspace === 'spreadsheet' ? (
          <SpreadsheetWorkspace
            historyRevision={historyRevision}
            mode={location.mode}
            onModeChange={(mode) => pushLocation({ ...location, mode })}
            onRecover={() => replaceLocation({ workspace: 'spreadsheet', mode: 'uncontrolled' })}
            onReload={onReload}
            restoredFromHistory={restoredFromHistory}
          />
        ) : null}
      </section>
      <section
        id="workspace-panel-output"
        className={styles.workspacePanel}
        role="tabpanel"
        aria-labelledby="workspace-tab-output"
        hidden={location.workspace !== 'output'}
      >
        {location.workspace === 'output' ? <OutputStudio embedded /> : null}
      </section>
    </main>
  );
}
