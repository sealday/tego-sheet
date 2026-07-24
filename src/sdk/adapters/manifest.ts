import type { AdapterKind, AdapterManifest } from './types';
import { adapterDiagnostic, AdapterSdkError } from './diagnostics';
import { JsonSnapshotError, snapshotStringList } from './json-safe';

export const ADAPTER_KINDS = Object.freeze([
  'workbook-reader',
  'workbook-writer',
  'resource-resolver',
  'output',
  'chart-renderer',
  'formula-function-provider',
  'solver',
  'persistence',
  'collaboration',
  'permission',
  'comments',
  'version-history',
  'ai-command',
] as const satisfies readonly AdapterKind[]);

const kindSet = new Set<string>(ADAPTER_KINDS);
const capabilityPattern = /^[a-z][a-z0-9]*(?:[.:-][a-z0-9]+)*$/;

/** Validates public fields and captures one immutable, getter-safe manifest snapshot. */
export function snapshotAdapterManifest<K extends AdapterKind>(
  manifest: AdapterManifest<K>,
): Readonly<AdapterManifest<K>> {
  let candidate: Readonly<Record<string, unknown>>;
  let id: unknown;
  let apiVersion: unknown;
  let kind: unknown;
  let environments: unknown;
  let execution: unknown;
  let priority: unknown;
  let capabilities: unknown;
  let formats: unknown;
  try {
    if ((typeof manifest !== 'object' && typeof manifest !== 'function') || manifest === null) {
      throw new TypeError('Manifest must be an object');
    }
    candidate = manifest as unknown as Readonly<Record<string, unknown>>;
    id = candidate.id;
    apiVersion = candidate.apiVersion;
    kind = candidate.kind;
    environments = candidate.environments;
    execution = candidate.execution;
    priority = candidate.priority;
    capabilities = candidate.capabilities;
    formats = candidate.formats;
  } catch (cause) {
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_MANIFEST_INVALID',
        'validate',
        'Adapter manifest could not be read safely',
        { cause },
      ),
    ]);
  }

  let capabilitySnapshot: readonly string[] = [];
  let formatSnapshot: readonly string[] | undefined;
  let environmentSnapshot: readonly string[] = [];
  try {
    capabilitySnapshot = snapshotStringList(
      capabilities,
      'manifest.capabilities',
      (entry) => capabilityPattern.test(entry),
      false,
    );
    environmentSnapshot = snapshotStringList(
      environments,
      'manifest.environments',
      (entry) => entry === 'browser' || entry === 'worker' || entry === 'node',
      false,
    );
    formatSnapshot =
      formats === undefined
        ? undefined
        : snapshotStringList(formats, 'manifest.formats', (entry) => entry.length > 0, false);
  } catch (cause) {
    if (!(cause instanceof JsonSnapshotError)) throw cause;
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_MANIFEST_INVALID',
        'validate',
        cause.message,
        typeof id === 'string' ? { manifest: { id }, cause } : { cause },
      ),
    ]);
  }
  if (
    typeof id !== 'string' ||
    typeof apiVersion !== 'string' ||
    typeof kind !== 'string' ||
    !kindSet.has(kind) ||
    (execution !== 'trusted-main' && execution !== 'isolated-worker') ||
    typeof priority !== 'number' ||
    !Number.isSafeInteger(priority)
  ) {
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_MANIFEST_INVALID',
        'validate',
        'Adapter manifest has invalid kind, trust, priority, capabilities, or formats',
        typeof id === 'string' ? { manifest: { id } } : {},
      ),
    ]);
  }

  return Object.freeze({
    id,
    apiVersion: apiVersion as AdapterManifest<K>['apiVersion'],
    kind: kind as K,
    environments: environmentSnapshot as AdapterManifest<K>['environments'],
    execution,
    priority,
    capabilities: capabilitySnapshot,
    ...(formatSnapshot === undefined ? {} : { formats: formatSnapshot }),
  });
}
