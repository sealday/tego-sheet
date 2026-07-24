/** Declared execution boundary for one public extension or adapter. */
export type ExtensionExecution = 'trusted-main' | 'isolated-worker';

/** Immutable capability allow-list granted to one operation scope. */
export interface CapabilityGrant {
  /** Stable capabilities authorized by the host. */
  readonly capabilities: readonly string[];
  /** Tests whether one exact capability was granted. */
  allows(capability: string): boolean;
}

const capabilityPattern = /^[a-z][a-z0-9]*(?:[.:-][a-z0-9]+)*$/;

/** Creates a deduplicated immutable capability grant. */
export function createCapabilityGrant(capabilities: readonly string[]): CapabilityGrant {
  const snapshot = Object.freeze(
    [
      ...snapshotStringList(
        capabilities,
        'capabilities',
        (entry) => capabilityPattern.test(entry),
        true,
      ),
    ].sort(),
  );
  const allowed = new Set(snapshot);
  return Object.freeze({
    capabilities: snapshot,
    allows: (capability: string): boolean => allowed.has(capability),
  });
}

/** Bounded resources available to one adapter scope. */
export interface AdapterScopeLimits {
  /** Maximum simultaneous adapter invocations. */
  readonly maxConcurrentInvocations: number;
  /** Maximum elapsed duration for one invocation. */
  readonly maxDurationMs: number;
  /** Maximum JSON-encoded invocation input. */
  readonly maxInputBytes: number;
  /** Maximum JSON-encoded invocation result. */
  readonly maxOutputBytes: number;
}

/** Conservative default limits for host adapter invocations. */
export const DEFAULT_ADAPTER_SCOPE_LIMITS: Readonly<AdapterScopeLimits> = Object.freeze({
  maxConcurrentInvocations: 4,
  maxDurationMs: 30_000,
  maxInputBytes: 1024 * 1024,
  maxOutputBytes: 4 * 1024 * 1024,
});
import { snapshotStringList } from '../adapters/json-safe';
