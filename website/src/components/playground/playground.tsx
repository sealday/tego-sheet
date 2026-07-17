import { TegoSheet, type TegoSheetHandle, type WorkbookData } from 'tego-sheet';
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
  type ReactElement,
} from 'react';
import { PLAYGROUND_LOCALES, PLAYGROUND_PRESETS, createFixture } from './playground-fixtures';
import {
  appendPlaygroundEvent,
  parsePlaygroundMode,
  type PlaygroundCallbackName,
  type PlaygroundEvent,
  type PlaygroundMode,
} from './playground-model';
import { PlaygroundErrorBoundary } from './playground-error-boundary';
import styles from './playground.module.css';

type SheetProps = ComponentProps<typeof TegoSheet>;
type SheetCallbacks = Pick<
  SheetProps,
  'onChange' | 'onActiveSheetChange' | 'onSelectionChange' | 'onCellEdit' | 'onPaste' | 'onError'
>;
type ToolbarRenderer = Exclude<SheetProps['toolbar'], 'default' | false | undefined>;
type SheetTabsRenderer = Exclude<SheetProps['sheetTabs'], 'default' | false | undefined>;

const LOCALES = { en, 'zh-CN': zhCN, de, nl } as const;
type LocaleId = keyof typeof LOCALES;

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

function relativeUrlForMode(mode: PlaygroundMode): string {
  const url = new URL(window.location.href);
  url.searchParams.set('mode', mode);
  return `${url.pathname}${url.search}${url.hash}`;
}

function modeFromLocation(): PlaygroundMode {
  return parsePlaygroundMode(new URLSearchParams(window.location.search).get('mode'));
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

interface PresetSheetProps {
  readonly fixture: WorkbookData;
  readonly locale: (typeof LOCALES)[LocaleId];
  readonly callbacks: SheetCallbacks;
  readonly sheetRef: React.RefObject<TegoSheetHandle | null>;
}

function ControlledSheet({ fixture, locale, callbacks, sheetRef }: PresetSheetProps): ReactElement {
  const [value, setValue] = useState<WorkbookData>(fixture);
  const onChange: NonNullable<SheetCallbacks['onChange']> = (nextValue, change) => {
    setValue(nextValue);
    callbacks.onChange?.(nextValue, change);
  };

  return (
    <TegoSheet ref={sheetRef} value={value} locale={locale} {...callbacks} onChange={onChange} />
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
      <button type="button" onClick={() => props.add('Demo sheet')}>
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
      defaultValue={fixture}
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
  const [events, setEvents] = useState<readonly PlaygroundEvent[]>([]);
  const [snapshot, setSnapshot] = useState<WorkbookData>(fixture);
  const [localeId, setLocaleId] = useState<LocaleId>('en');

  const record = useCallback((callback: PlaygroundCallbackName, payload: unknown): void => {
    setEvents((current) =>
      appendPlaygroundEvent(current, {
        sequence: ++sequence.current,
        callback,
        payload: toPublicJson(payload),
      }),
    );
  }, []);

  const callbacks = useMemo<SheetCallbacks>(
    () => ({
      onChange: (value, change) => {
        setSnapshot(value);
        record('onChange', change);
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
  const copySnapshot = async (): Promise<void> => {
    try {
      await navigator.clipboard.writeText(formattedSnapshot);
      setStatus('Workbook JSON copied');
    } catch {
      setStatus('Could not copy workbook JSON');
    }
  };
  const refreshSnapshot = (): void => {
    const currentValue = sheetRef.current?.getValue();
    if (currentValue) setSnapshot(currentValue);
    setStatus('Workbook JSON refreshed from TegoSheetHandle.getValue()');
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
            <h3>Workbook JSON</h3>
            <div className={styles.buttonRow}>
              <button type="button" onClick={refreshSnapshot}>
                Refresh JSON
              </button>
              <button type="button" onClick={() => void copySnapshot()}>
                Copy JSON
              </button>
            </div>
          </div>
          <pre className={styles.json} aria-label="Workbook JSON" tabIndex={0}>
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

export function Playground({ onReload = reloadWindow }: PlaygroundProps = {}): ReactElement {
  const [initialMode] = useState<PlaygroundMode>(modeFromLocation);
  const [mode, setMode] = useState<PlaygroundMode>(initialMode);
  const [resetRevision, setResetRevision] = useState(0);
  const [status, setStatus] = useState('');

  const replaceMode = useCallback((nextMode: PlaygroundMode): void => {
    window.history.replaceState(window.history.state, '', relativeUrlForMode(nextMode));
  }, []);

  useEffect(() => {
    const rawMode = new URLSearchParams(window.location.search).get('mode');
    if (rawMode !== initialMode) replaceMode(initialMode);
  }, [initialMode, replaceMode]);

  useEffect(() => {
    const onPopState = (): void => {
      const nextMode = modeFromLocation();
      const rawMode = new URLSearchParams(window.location.search).get('mode');
      if (rawMode !== nextMode) replaceMode(nextMode);
      setMode(nextMode);
      setResetRevision((revision) => revision + 1);
      setStatus(`${PLAYGROUND_PRESETS[nextMode].label} restored from browser history`);
    };
    window.addEventListener('popstate', onPopState);
    return () => window.removeEventListener('popstate', onPopState);
  }, [replaceMode]);

  const selectMode = (nextMode: PlaygroundMode): void => {
    if (nextMode === mode) return;
    window.history.pushState(window.history.state, '', relativeUrlForMode(nextMode));
    setMode(nextMode);
    setResetRevision(0);
    setStatus(`${PLAYGROUND_PRESETS[nextMode].label} selected`);
  };

  const resetMode = (): void => {
    setResetRevision((revision) => revision + 1);
    setStatus(`${PLAYGROUND_PRESETS[mode].label} reset`);
  };

  const recoverFromError = (): void => {
    replaceMode('uncontrolled');
    setMode('uncontrolled');
    setResetRevision((revision) => revision + 1);
    setStatus('Playground reset to Uncontrolled');
  };

  const presetKey = `${mode}:${resetRevision}`;

  return (
    <main className={styles.playground}>
      <header className={styles.header}>
        <div>
          <p className={styles.eyebrow}>Live public API examples</p>
          <h1>Playground</h1>
          <p>Switch presets without leaving the page, then inspect workbook data and callbacks.</p>
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
        {status}
      </p>
      <PlaygroundErrorBoundary onReset={recoverFromError} onReload={onReload}>
        <PresetSession
          key={presetKey}
          mode={mode}
          presetKey={presetKey}
          setStatus={setStatus}
          onReset={resetMode}
        />
      </PlaygroundErrorBoundary>
    </main>
  );
}
