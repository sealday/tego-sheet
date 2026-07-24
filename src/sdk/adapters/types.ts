import type { JsonValue } from '../../core/types/json';
import type { Diagnostic } from '../../document/diagnostics';
import type { KernelEnvironment } from '../../extensions/kernel/manifest';
import type { WorkbookReader, WorkbookWriter } from '../../interchange/contracts';
import type { ResourceResolver } from '../../template/resources';
import type { AdapterScopeLimits, CapabilityGrant, ExtensionExecution } from '../trust';

/** Public adapter API compatibility version. */
export type AdapterApiVersion = `${number}.${number}`;

/** Runtime environment in which an adapter may be resolved. */
export type AdapterEnvironment = KernelEnvironment;

/** Request delivered to a trusted adapter implementation. */
export interface AdapterInvocationRequest {
  /** Exact capability selected by the caller. */
  readonly capability: string;
  /** Capability-specific input owned by the caller. */
  readonly input: unknown;
}

/** Restricted context available during a trusted-main invocation. */
export interface AdapterInvocationContext {
  /** Optional document identity for operation scoping, never a controller reference. */
  readonly documentId?: string;
  /** Invocation-local cancellation signal. */
  readonly signal: AbortSignal;
}

/** Minimal callable protocol used by generic adapter kinds. */
export interface CallableAdapter {
  /** Executes one declared capability through the scope boundary. */
  invoke(
    request: AdapterInvocationRequest,
    context: AdapterInvocationContext,
  ): unknown | Promise<unknown>;
}

/** Pluggable optimization adapter entry point. */
export interface SolverAdapter extends CallableAdapter {
  readonly adapterKind?: 'solver';
}

/** Host-owned persistence adapter entry point. */
export interface PersistenceAdapter extends CallableAdapter {
  readonly adapterKind?: 'persistence';
}

/** Host-owned collaboration adapter entry point. */
export interface CollaborationAdapter extends CallableAdapter {
  readonly adapterKind?: 'collaboration';
}

/** Host-owned permission adapter entry point. */
export interface PermissionAdapter extends CallableAdapter {
  readonly adapterKind?: 'permission';
}

/** Host-owned comments adapter entry point. */
export interface CommentsAdapter extends CallableAdapter {
  readonly adapterKind?: 'comments';
}

/** Host-owned durable version-history adapter entry point. */
export interface VersionHistoryAdapter extends CallableAdapter {
  readonly adapterKind?: 'version-history';
}

/** Host-owned AI command proposal adapter entry point. */
export interface AiCommandAdapter extends CallableAdapter {
  readonly adapterKind?: 'ai-command';
}

/** Optional renderer-neutral chart adapter entry point. */
export interface ChartRendererAdapter extends CallableAdapter {
  readonly adapterKind?: 'chart-renderer';
}

/** Explicit formula-function provider adapter entry point. */
export interface FormulaFunctionProviderAdapter extends CallableAdapter {
  readonly adapterKind?: 'formula-function-provider';
}

/** Generic document output adapter entry point. */
export interface OutputAdapter extends CallableAdapter {
  readonly adapterKind?: 'output';
}

/**
 * Open, kind-indexed public adapter implementation map.
 *
 * Future packages may augment this interface with declaration merging.
 */
export interface AdapterByKind {
  readonly 'workbook-reader': WorkbookReader;
  readonly 'workbook-writer': WorkbookWriter;
  readonly 'resource-resolver': ResourceResolver;
  readonly output: OutputAdapter;
  readonly 'chart-renderer': ChartRendererAdapter;
  readonly 'formula-function-provider': FormulaFunctionProviderAdapter;
  readonly solver: SolverAdapter;
  readonly persistence: PersistenceAdapter;
  readonly collaboration: CollaborationAdapter;
  readonly permission: PermissionAdapter;
  readonly comments: CommentsAdapter;
  readonly 'version-history': VersionHistoryAdapter;
  readonly 'ai-command': AiCommandAdapter;
}

/** Registered public adapter discriminator. */
export type AdapterKind = keyof AdapterByKind & string;

/** Immutable public manifest retained by the registry. */
export interface AdapterManifest<K extends AdapterKind = AdapterKind> {
  /** Stable lowercase adapter identity. */
  readonly id: string;
  /** Public API compatibility version. */
  readonly apiVersion: AdapterApiVersion;
  /** Kind-specific protocol implemented by the adapter. */
  readonly kind: K;
  /** Supported runtime environments. */
  readonly environments: readonly AdapterEnvironment[];
  /** Explicit same-realm or worker-isolated execution declaration. */
  readonly execution: ExtensionExecution;
  /** Stable list ordering only; it never resolves ambiguity by itself. */
  readonly priority: number;
  /** Exact capabilities which may be granted to invocation scopes. */
  readonly capabilities: readonly string[];
  /** Optional normalized formats supported by the adapter. */
  readonly formats?: readonly string[];
}

/** Initialization context supplied through the existing F5 kernel lifecycle. */
export interface AdapterInitializationContext {
  /** Host runtime environment. */
  readonly environment: AdapterEnvironment;
  /** Registration lifetime cancellation signal. */
  readonly signal: AbortSignal;
  /** Safe diagnostic publisher owned by the host registry. */
  readonly diagnostics: (diagnostic: Diagnostic) => void;
}

/** One public adapter and its F5-backed lifecycle hooks. */
export interface AdapterRegistration<K extends AdapterKind = AdapterKind> {
  /** Immutable registration metadata. */
  readonly manifest: AdapterManifest<K>;
  /** Kind-specific implementation retained by the F5 kernel. */
  readonly implementation: AdapterByKind[K];
  /** Optional initialization before publication. */
  readonly initialize?: (context: AdapterInitializationContext) => void | Promise<void>;
  /** Optional cleanup after the adapter is unpublished. */
  readonly dispose?: () => void | Promise<void>;
}

/** Public query used to list adapter manifests. */
export interface AdapterQuery {
  readonly kind?: AdapterKind;
  readonly capability?: string;
  readonly format?: string;
}

/** Public query used to resolve exactly one adapter. */
export interface AdapterResolutionQuery {
  readonly id?: string;
  readonly capability?: string;
  readonly format?: string;
}

/** Why one adapter was selected. */
export type AdapterResolutionReason = 'explicit-id' | 'single-match' | 'configured-default';

/** Stable resolved adapter snapshot. */
export interface AdapterResolution<K extends AdapterKind = AdapterKind> {
  readonly manifest: AdapterManifest<K>;
  readonly implementation: AdapterByKind[K] | IsolatedWorkerAdapterReference;
  readonly reason: AdapterResolutionReason;
}

/** Host-owned result schema predicate for one invocation. */
export type AdapterResultValidator<Result> =
  | ((value: unknown) => value is Result)
  | ((value: unknown) => boolean);

/** One bounded invocation request. */
export interface ScopedAdapterInvocation<Result> {
  readonly capability: string;
  readonly input: unknown;
  readonly validateResult: AdapterResultValidator<Result>;
}

/** Options used to create one document or operation scope. */
export interface AdapterScopeOptions {
  readonly documentId?: string;
  readonly signal: AbortSignal;
  readonly grant: CapabilityGrant;
  readonly limits?: Partial<AdapterScopeLimits>;
}

/** Isolated-worker request containing only schema-cloneable public data. */
export interface IsolatedWorkerInvocation {
  readonly adapterId: string;
  readonly kind: AdapterKind;
  readonly capability: string;
  readonly input: unknown;
  readonly documentId?: string;
}

/** Host transport which owns the actual Worker and structured-clone protocol. */
export interface IsolatedWorkerTransport {
  invoke(request: IsolatedWorkerInvocation, signal: AbortSignal): Promise<unknown>;
}

/** Opaque public reference returned instead of same-realm code for isolated adapters. */
export interface IsolatedWorkerAdapterReference {
  readonly execution: 'isolated-worker';
  readonly adapterId: string;
}

/** Operation-scoped invocation and cleanup surface. */
export interface AdapterScope {
  invoke<K extends AdapterKind, Result>(
    resolution: AdapterResolution<K>,
    invocation: ScopedAdapterInvocation<Result>,
  ): Promise<Result>;
  dispose(): Promise<readonly AdapterDiagnostic[]>;
}

/** Stable public adapter diagnostic codes. */
export type AdapterDiagnosticCode =
  | 'ADAPTER_MANIFEST_INVALID'
  | 'ADAPTER_DUPLICATE_ID'
  | 'ADAPTER_VERSION_UNSUPPORTED'
  | 'ADAPTER_ENVIRONMENT_UNSUPPORTED'
  | 'ADAPTER_NOT_FOUND'
  | 'ADAPTER_AMBIGUOUS'
  | 'ADAPTER_INITIALIZATION_FAILED'
  | 'ADAPTER_INVOCATION_FAILED'
  | 'ADAPTER_INPUT_INVALID'
  | 'ADAPTER_RESULT_INVALID'
  | 'ADAPTER_DISPOSE_FAILED'
  | 'ADAPTER_LIMIT_EXCEEDED'
  | 'ADAPTER_INVOCATION_TIMEOUT'
  | 'ADAPTER_INVOCATION_ABORTED'
  | 'ADAPTER_SCOPE_DISPOSED'
  | 'ADAPTER_REGISTRY_DISPOSED'
  | 'ADAPTER_OPTIONS_INVALID'
  | 'CAPABILITY_DENIED';

/** Diagnostic emitted by the public adapter SDK. */
export interface AdapterDiagnostic extends Diagnostic {
  readonly code: AdapterDiagnosticCode;
}

/** Public registry construction options. */
export interface AdapterRegistryOptions {
  readonly apiVersion: AdapterApiVersion;
  readonly environment: AdapterEnvironment;
  readonly defaults?: Readonly<Partial<Record<AdapterKind, string>>>;
  readonly diagnostics?: (diagnostic: AdapterDiagnostic) => void;
  readonly isolatedWorkerTransport?: IsolatedWorkerTransport;
}

/** Public F5-backed adapter registry. */
export interface AdapterRegistry {
  register<K extends AdapterKind>(
    registration: AdapterRegistration<K>,
  ): Promise<() => Promise<readonly AdapterDiagnostic[]>>;
  list(query?: AdapterQuery): readonly AdapterManifest[];
  resolve<K extends AdapterKind>(kind: K, query?: AdapterResolutionQuery): AdapterResolution<K>;
  createScope(options: AdapterScopeOptions): AdapterScope;
  dispose(): Promise<readonly AdapterDiagnostic[]>;
}

/** JSON-safe diagnostic details helper type. */
export type AdapterDiagnosticDetails = JsonValue;
