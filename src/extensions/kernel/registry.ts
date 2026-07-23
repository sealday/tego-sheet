import type { Diagnostic } from '../../document/diagnostics';
import type { KernelCapabilities, KernelExtensionKind, KernelRegistration } from './capabilities';
import {
  assertSupportedApiVersion,
  extensionDiagnostic,
  ExtensionKernelError,
  validateAndSnapshotManifest,
} from './manifest';
import type {
  ApiVersion,
  ExtensionDiagnostic,
  ExtensionManifest,
  KernelEnvironment,
} from './manifest';

/** Internal typed adapter registry. */
export interface AdapterRegistryKernel {
  register<K extends KernelExtensionKind>(
    registration: KernelRegistration<K>,
  ): Promise<() => Promise<readonly Diagnostic[]>>;
  list(kind?: KernelExtensionKind): readonly ExtensionManifest[];
  resolve<K extends KernelExtensionKind>(
    kind: K,
    query: { readonly id?: string; readonly environment: KernelEnvironment },
  ): KernelCapabilities[K];
  dispose(): Promise<readonly Diagnostic[]>;
}

export interface AdapterRegistryKernelOptions {
  readonly apiVersion: ApiVersion;
  readonly environment: KernelEnvironment;
  readonly diagnostics?: (diagnostic: Diagnostic) => void;
}

interface RegistrationRecord {
  readonly key: string;
  readonly manifest: Readonly<ExtensionManifest>;
  readonly implementation: KernelCapabilities[KernelExtensionKind];
  readonly controller: AbortController;
  readonly dispose?: () => void | Promise<void>;
  release?: Promise<readonly ExtensionDiagnostic[]>;
}

interface PendingRegistration {
  readonly key: string;
  readonly controller: AbortController;
  readonly completion: Promise<readonly ExtensionDiagnostic[]>;
}

function keyOf(kind: string, id: string): string {
  return `${kind}\u0000${id}`;
}

function compareRecords(left: RegistrationRecord, right: RegistrationRecord): number {
  return (
    left.manifest.kind.localeCompare(right.manifest.kind) ||
    left.manifest.id.localeCompare(right.manifest.id)
  );
}

function manifestForError(kind: KernelExtensionKind, id: string): ExtensionManifest {
  return { id, kind, apiVersion: '0.0', environments: [] };
}

/** Creates an isolated registry; no process-global registrations are mutated. */
export function createAdapterRegistryKernel(
  options: AdapterRegistryKernelOptions,
): AdapterRegistryKernel {
  const records = new Map<string, RegistrationRecord>();
  const pending = new Map<string, PendingRegistration>();
  let listCache: readonly ExtensionManifest[] | undefined;
  let disposed = false;
  let disposePromise: Promise<readonly Diagnostic[]> | undefined;
  let disposeComplete = false;

  const publishDiagnostic = (diagnostic: ExtensionDiagnostic): void => {
    options.diagnostics?.(diagnostic);
  };

  const fail = (diagnostic: ExtensionDiagnostic): never => {
    publishDiagnostic(diagnostic);
    throw new ExtensionKernelError([diagnostic]);
  };

  const release = (record: RegistrationRecord): Promise<readonly ExtensionDiagnostic[]> => {
    if (record.release !== undefined) return record.release;

    records.delete(record.key);
    listCache = undefined;
    record.controller.abort();
    record.release = (async () => {
      if (record.dispose === undefined) return Object.freeze([]);
      try {
        await record.dispose();
        return Object.freeze([]);
      } catch (cause) {
        const diagnostic = extensionDiagnostic(
          'EXTENSION_DISPOSE_FAILED',
          'dispose',
          `Failed to dispose extension ${record.manifest.kind}/${record.manifest.id}`,
          record.manifest,
          undefined,
          cause,
        );
        publishDiagnostic(diagnostic);
        return Object.freeze([diagnostic]);
      }
    })();
    return record.release;
  };

  return {
    async register<K extends KernelExtensionKind>(
      registration: KernelRegistration<K>,
    ): Promise<() => Promise<readonly Diagnostic[]>> {
      const manifest = validateAndSnapshotManifest(registration.manifest);
      assertSupportedApiVersion(manifest, options.apiVersion);
      const key = keyOf(manifest.kind, manifest.id);

      if (disposed) {
        fail(
          extensionDiagnostic(
            'EXTENSION_REGISTRY_DISPOSED',
            'validate',
            'Cannot register an extension after the registry has been disposed',
            manifest,
          ),
        );
      }
      if (records.has(key) || pending.has(key)) {
        fail(
          extensionDiagnostic(
            'EXTENSION_DUPLICATE_ID',
            'validate',
            `Extension ${manifest.kind}/${manifest.id} is already registered`,
            manifest,
          ),
        );
      }

      const controller = new AbortController();
      let completePending!: (diagnostics: readonly ExtensionDiagnostic[]) => void;
      const completion = new Promise<readonly ExtensionDiagnostic[]>((resolve) => {
        completePending = resolve;
      });
      pending.set(key, { key, controller, completion });
      const record: RegistrationRecord = {
        key,
        manifest,
        implementation: registration.implementation,
        controller,
        ...(registration.dispose === undefined ? {} : { dispose: registration.dispose }),
      };

      try {
        await registration.initialize?.({
          environment: options.environment,
          signal: controller.signal,
          diagnostics: options.diagnostics ?? (() => {}),
        });
      } catch (cause) {
        controller.abort();
        const initializationDiagnostic = extensionDiagnostic(
          'EXTENSION_INITIALIZE_FAILED',
          'execute',
          `Failed to initialize extension ${manifest.kind}/${manifest.id}`,
          manifest,
          undefined,
          cause,
        );
        publishDiagnostic(initializationDiagnostic);
        const cleanupDiagnostics = await release(record);
        pending.delete(key);
        completePending(cleanupDiagnostics);
        throw new ExtensionKernelError([initializationDiagnostic, ...cleanupDiagnostics]);
      }

      if (disposed) {
        const cleanupDiagnostics = await release(record);
        pending.delete(key);
        completePending(cleanupDiagnostics);
        const disposedDiagnostic = extensionDiagnostic(
          'EXTENSION_REGISTRY_DISPOSED',
          'validate',
          'The registry was disposed while the extension initialized',
          manifest,
        );
        publishDiagnostic(disposedDiagnostic);
        throw new ExtensionKernelError([disposedDiagnostic, ...cleanupDiagnostics]);
      }
      records.set(key, record);
      listCache = undefined;
      pending.delete(key);
      completePending(Object.freeze([]));

      return () => release(record);
    },

    list(kind?: KernelExtensionKind): readonly ExtensionManifest[] {
      if (listCache === undefined) {
        listCache = Object.freeze(
          [...records.values()].sort(compareRecords).map(({ manifest }) => manifest),
        );
      }
      if (kind === undefined) return listCache;
      return Object.freeze(listCache.filter((manifest) => manifest.kind === kind));
    },

    resolve<K extends KernelExtensionKind>(
      kind: K,
      query: { readonly id?: string; readonly environment: KernelEnvironment },
    ): KernelCapabilities[K] {
      if (query.id !== undefined) {
        const record = records.get(keyOf(kind, query.id));
        if (record === undefined) {
          return fail(
            extensionDiagnostic(
              'EXTENSION_NOT_FOUND',
              'validate',
              `Extension ${kind}/${query.id} is not registered`,
              manifestForError(kind, query.id),
            ),
          );
        }
        if (!record.manifest.environments.includes(query.environment)) {
          fail(
            extensionDiagnostic(
              'EXTENSION_ENVIRONMENT_UNSUPPORTED',
              'validate',
              `Extension ${kind}/${query.id} does not support ${query.environment}`,
              record.manifest,
              { environment: query.environment },
            ),
          );
        }
        return record.implementation as KernelCapabilities[K];
      }

      const sameKind = [...records.values()]
        .filter(({ manifest }) => manifest.kind === kind)
        .sort(compareRecords);
      const candidates = sameKind.filter(({ manifest }) =>
        manifest.environments.includes(query.environment),
      );
      if (candidates.length === 0) {
        const code =
          sameKind.length === 0 ? 'EXTENSION_NOT_FOUND' : 'EXTENSION_ENVIRONMENT_UNSUPPORTED';
        fail(
          extensionDiagnostic(
            code,
            'validate',
            `No ${kind} extension supports ${query.environment}`,
            undefined,
            { environment: query.environment, kind },
          ),
        );
      }
      if (candidates.length > 1) {
        fail(
          extensionDiagnostic(
            'EXTENSION_AMBIGUOUS',
            'validate',
            `More than one ${kind} extension supports ${query.environment}; specify an ID`,
            undefined,
            { candidates: candidates.map(({ manifest }) => manifest.id), kind },
          ),
        );
      }
      return candidates[0]!.implementation as KernelCapabilities[K];
    },

    dispose(): Promise<readonly Diagnostic[]> {
      if (disposePromise !== undefined) {
        return disposeComplete ? Promise.resolve(Object.freeze([])) : disposePromise;
      }
      disposed = true;
      disposePromise = (async () => {
        const diagnostics: ExtensionDiagnostic[] = [];
        const initializing = [...pending.values()].sort((left, right) =>
          left.key.localeCompare(right.key),
        );
        for (const registration of initializing) registration.controller.abort();
        for (const record of [...records.values()].sort(compareRecords)) {
          diagnostics.push(...(await release(record)));
        }
        for (const registration of initializing) {
          diagnostics.push(...(await registration.completion));
        }
        disposeComplete = true;
        return Object.freeze(diagnostics);
      })();
      return disposePromise;
    },
  };
}
