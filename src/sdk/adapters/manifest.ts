import type { AdapterKind, AdapterManifest } from './types';
import { adapterDiagnostic, AdapterSdkError } from './diagnostics';

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

  const validCapabilities =
    Array.isArray(capabilities) &&
    capabilities.length > 0 &&
    capabilities.every(
      (capability) => typeof capability === 'string' && capabilityPattern.test(capability),
    ) &&
    new Set(capabilities).size === capabilities.length;
  const validFormats =
    formats === undefined ||
    (Array.isArray(formats) &&
      formats.length > 0 &&
      formats.every((format) => typeof format === 'string' && format.length > 0) &&
      new Set(formats).size === formats.length);
  if (
    typeof id !== 'string' ||
    typeof apiVersion !== 'string' ||
    typeof kind !== 'string' ||
    !kindSet.has(kind) ||
    !Array.isArray(environments) ||
    (execution !== 'trusted-main' && execution !== 'isolated-worker') ||
    typeof priority !== 'number' ||
    !Number.isSafeInteger(priority) ||
    !validCapabilities ||
    !validFormats
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

  const environmentSnapshot = environments as readonly AdapterManifest<K>['environments'][number][];
  const capabilitySnapshot = capabilities as readonly string[];
  const formatSnapshot = formats as readonly string[] | undefined;
  return Object.freeze({
    id,
    apiVersion: apiVersion as AdapterManifest<K>['apiVersion'],
    kind: kind as K,
    environments: Object.freeze([...environmentSnapshot]),
    execution,
    priority,
    capabilities: Object.freeze([...capabilitySnapshot]),
    ...(formatSnapshot === undefined ? {} : { formats: Object.freeze([...formatSnapshot]) }),
  });
}
