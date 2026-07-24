import type { JsonValue } from '../../core/types/json';
import type { CellInput } from '../../document';
import type { AdapterApiVersion, AdapterEnvironment } from '../adapters';
import type { ExtensionExecution } from '../trust';

/** Persisted custom-cell input already supported by Workbook 2.0. */
export type CustomCellInput = Extract<CellInput, { readonly type: 'custom' }>;

/** Minimal manifest for one public custom-cell protocol implementation. */
export interface CellTypeManifest {
  readonly id: string;
  readonly apiVersion: AdapterApiVersion;
  readonly execution: ExtensionExecution;
  readonly environments: readonly AdapterEnvironment[];
  readonly capabilities: readonly string[];
}

export type CellDeserializeResult<Value> =
  | { readonly ok: true; readonly value: Value }
  | { readonly ok: false; readonly message: string };

/** Versioned, renderer-neutral custom-cell protocol. */
export interface CellTypePlugin<Value = unknown, Serialized extends JsonValue = JsonValue> {
  readonly manifest: CellTypeManifest;
  readonly type: string;
  readonly schemaVersion: number;
  validate(value: unknown): value is Value;
  deserialize(value: unknown, storedVersion: number): CellDeserializeResult<Value>;
  serialize(value: Value): Serialized;
  format(value: Value, context: CellFormatContext): string;
  print?: (value: Value, context: CellFormatContext) => string;
  accessibility?: (value: Value, context: CellFormatContext) => string;
}

export interface CellFormatContext {
  readonly locale: string;
}

export interface CellTypeRegistryOptions {
  readonly apiVersion: AdapterApiVersion;
  readonly environment: AdapterEnvironment;
}

export interface CellTypeRegistry {
  register<Value, Serialized extends JsonValue>(
    plugin: CellTypePlugin<Value, Serialized>,
  ): () => void;
  get(type: string): CellTypePlugin | undefined;
  list(): readonly CellTypePlugin[];
}

export type CellExtensionDiagnosticCode =
  | 'CELL_PLUGIN_UNAVAILABLE'
  | 'CELL_PLUGIN_VERSION_UNSUPPORTED'
  | 'CELL_VALUE_INVALID'
  | 'CELL_VALUE_MIGRATION_FAILED'
  | 'CELL_RENDER_FAILED'
  | 'CELL_EDITOR_FAILED'
  | 'CELL_PRINT_FALLBACK';

export interface CellExtensionDiagnostic {
  readonly code: CellExtensionDiagnosticCode;
  readonly message: string;
}

export type ResolvedCustomCell =
  | {
      readonly status: 'resolved';
      readonly value: unknown;
      readonly formattedText: string;
      readonly accessibilityLabel: string;
      readonly printText: string;
      readonly diagnostics: readonly CellExtensionDiagnostic[];
    }
  | {
      readonly status: 'fallback';
      readonly serializedInput: CustomCellInput;
      readonly formattedText: string;
      readonly accessibilityLabel: string;
      readonly printText: string;
      readonly diagnostics: readonly CellExtensionDiagnostic[];
    };

const typePattern = /^[a-z][a-z0-9]*(?:[.-][a-z0-9]+)+$/;
const manifestIdPattern = /^[a-z][a-z0-9]*(?:[.-][a-z0-9]+)*$/;
const maximumCellValueBytes = 65_536;

function major(version: AdapterApiVersion): string {
  return version.split('.')[0] as string;
}

function canonicalJson(value: JsonValue): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`;
  const record = value as Readonly<Record<string, JsonValue>>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${canonicalJson(record[key]!)}`)
    .join(',')}}`;
}

function snapshotJson(value: JsonValue): JsonValue {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return Object.freeze(value.map(snapshotJson));
  const record = value as Readonly<Record<string, JsonValue>>;
  return Object.freeze(
    Object.fromEntries(Object.keys(record).map((key) => [key, snapshotJson(record[key]!)])),
  );
}

function snapshotInput(input: CustomCellInput): CustomCellInput {
  return Object.freeze({
    type: 'custom',
    cellType: input.cellType,
    schemaVersion: input.schemaVersion,
    value: snapshotJson(input.value),
  });
}

/** Creates an isolated custom-cell registry with duplicate and compatibility checks. */
export function createCellTypeRegistry(options: CellTypeRegistryOptions): CellTypeRegistry {
  const records = new Map<string, CellTypePlugin>();
  return Object.freeze({
    register<Value, Serialized extends JsonValue>(
      plugin: CellTypePlugin<Value, Serialized>,
    ): () => void {
      if (!typePattern.test(plugin.type)) throw new TypeError('Cell plugin type is invalid');
      if (!manifestIdPattern.test(plugin.manifest.id)) {
        throw new TypeError('Cell plugin manifest ID is invalid');
      }
      if (!Number.isSafeInteger(plugin.schemaVersion) || plugin.schemaVersion < 1) {
        throw new TypeError('Cell plugin schemaVersion must be a positive integer');
      }
      if (major(plugin.manifest.apiVersion) !== major(options.apiVersion)) {
        throw new TypeError(`Cell plugin ${plugin.type} API version is unsupported`);
      }
      if (!plugin.manifest.environments.includes(options.environment)) {
        throw new TypeError(`Cell plugin ${plugin.type} environment is unsupported`);
      }
      if (records.has(plugin.type)) {
        throw new TypeError(`Cell plugin type ${plugin.type} is already registered`);
      }
      const snapshot = Object.freeze({
        ...plugin,
        manifest: Object.freeze({
          ...plugin.manifest,
          environments: Object.freeze([...plugin.manifest.environments]),
          capabilities: Object.freeze([...plugin.manifest.capabilities]),
        }),
      }) as CellTypePlugin;
      records.set(plugin.type, snapshot);
      let active = true;
      return (): void => {
        if (!active) return;
        active = false;
        if (records.get(plugin.type) === snapshot) records.delete(plugin.type);
      };
    },
    get(type: string): CellTypePlugin | undefined {
      return records.get(type);
    },
    list(): readonly CellTypePlugin[] {
      return Object.freeze(
        Array.from(records.values()).sort((left, right) => left.type.localeCompare(right.type)),
      );
    },
  });
}

function fallback(input: CustomCellInput, diagnostic: CellExtensionDiagnostic): ResolvedCustomCell {
  const serializedInput = snapshotInput(input);
  const formattedText = canonicalJson(serializedInput.value);
  return Object.freeze({
    status: 'fallback',
    serializedInput,
    formattedText,
    accessibilityLabel: formattedText,
    printText: formattedText,
    diagnostics: Object.freeze([Object.freeze(diagnostic)]),
  });
}

/** Resolves one stored custom value or returns a lossless, plain-text fallback. */
export function resolveCustomCell(
  registry: CellTypeRegistry,
  input: CustomCellInput,
  context: CellFormatContext,
): ResolvedCustomCell {
  const serialized = canonicalJson(input.value);
  if (new TextEncoder().encode(serialized).byteLength > maximumCellValueBytes) {
    throw new RangeError(`Custom cell values cannot exceed ${maximumCellValueBytes} bytes`);
  }
  const plugin = registry.get(input.cellType);
  if (plugin === undefined) {
    return fallback(input, {
      code: 'CELL_PLUGIN_UNAVAILABLE',
      message: `Cell plugin ${input.cellType} is unavailable`,
    });
  }
  let decoded: CellDeserializeResult<unknown>;
  try {
    decoded = plugin.deserialize(snapshotJson(input.value), input.schemaVersion);
  } catch (cause) {
    return fallback(input, {
      code: 'CELL_VALUE_MIGRATION_FAILED',
      message: cause instanceof Error ? cause.message : `Cell plugin ${input.cellType} failed`,
    });
  }
  if (!decoded.ok) {
    return fallback(input, {
      code: 'CELL_VALUE_MIGRATION_FAILED',
      message: decoded.message,
    });
  }
  if (!plugin.validate(decoded.value)) {
    return fallback(input, {
      code: 'CELL_VALUE_INVALID',
      message: `Cell plugin ${input.cellType} rejected the stored value`,
    });
  }
  try {
    const formattedText = plugin.format(decoded.value, context);
    const accessibilityLabel = plugin.accessibility?.(decoded.value, context) ?? formattedText;
    const diagnostics: CellExtensionDiagnostic[] = [];
    const printText =
      plugin.print?.(decoded.value, context) ??
      (() => {
        diagnostics.push({
          code: 'CELL_PRINT_FALLBACK',
          message: `Cell plugin ${input.cellType} has no print renderer; formatted text was used`,
        });
        return formattedText;
      })();
    return Object.freeze({
      status: 'resolved',
      value: decoded.value,
      formattedText,
      accessibilityLabel,
      printText,
      diagnostics: Object.freeze(diagnostics.map((diagnostic) => Object.freeze(diagnostic))),
    });
  } catch (cause) {
    return fallback(input, {
      code: 'CELL_RENDER_FAILED',
      message: cause instanceof Error ? cause.message : `Cell plugin ${input.cellType} failed`,
    });
  }
}

export interface CreateCellEditorSessionOptions<Value> {
  readonly initialValue: Value;
  readonly signal: AbortSignal;
  readonly validate: (value: unknown) => value is Value;
  readonly commit: (value: Value) => void;
  readonly cancel: () => void;
}

export interface CellEditorSession<Value> {
  readonly initialValue: Value;
  readonly signal: AbortSignal;
  commit(value: Value): boolean;
  cancel(): boolean;
  dispose(): void;
}

/** Creates an exactly-once editor handoff which never writes the document directly. */
export function createCellEditorSession<Value>(
  options: CreateCellEditorSessionOptions<Value>,
): CellEditorSession<Value> {
  let settled = false;
  const onAbort = (): void => {
    if (settled) return;
    settled = true;
    options.cancel();
  };
  options.signal.addEventListener('abort', onAbort, { once: true });
  if (options.signal.aborted) onAbort();
  const settle = (kind: 'commit' | 'cancel', value?: Value): boolean => {
    if (settled || options.signal.aborted) return false;
    if (kind === 'commit' && !options.validate(value)) {
      throw new TypeError('Cell editor attempted to commit an invalid value');
    }
    settled = true;
    options.signal.removeEventListener('abort', onAbort);
    if (kind === 'commit') options.commit(value as Value);
    else options.cancel();
    return true;
  };
  return Object.freeze({
    initialValue: options.initialValue,
    signal: options.signal,
    commit: (value: Value): boolean => settle('commit', value),
    cancel: (): boolean => settle('cancel'),
    dispose(): void {
      options.signal.removeEventListener('abort', onAbort);
    },
  });
}
