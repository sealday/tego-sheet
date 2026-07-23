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

function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function compareRecords(left: RegistrationRecord, right: RegistrationRecord): number {
  return (
    compareAscii(left.manifest.kind, right.manifest.kind) ||
    compareAscii(left.manifest.id, right.manifest.id)
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
  const releasing = new Map<string, Promise<readonly ExtensionDiagnostic[]>>();
  let listCache: readonly ExtensionManifest[] | undefined;
  let disposed = false;
  let disposePromise: Promise<readonly Diagnostic[]> | undefined;
  let disposeComplete = false;

  const publishDiagnostic = (diagnostic: Diagnostic): void => {
    try {
      options.diagnostics?.(diagnostic);
    } catch {
      // Diagnostic observers are telemetry sinks and cannot own kernel control flow.
    }
  };

  const fail = (diagnostic: ExtensionDiagnostic): never => {
    publishDiagnostic(diagnostic);
    throw new ExtensionKernelError([diagnostic]);
  };

  const unpublish = (record: RegistrationRecord): void => {
    records.delete(record.key);
    listCache = undefined;
  };

  const detach = (record: RegistrationRecord): void => {
    unpublish(record);
    record.controller.abort();
  };

  const release = (record: RegistrationRecord): Promise<readonly ExtensionDiagnostic[]> => {
    if (record.release !== undefined) return record.release;

    let completeRelease!: (diagnostics: readonly ExtensionDiagnostic[]) => void;
    const task = new Promise<readonly ExtensionDiagnostic[]>((resolve) => {
      completeRelease = resolve;
    });
    record.release = task;
    releasing.set(record.key, task);
    void task.then(() => {
      if (releasing.get(record.key) === task) releasing.delete(record.key);
    });

    detach(record);
    void (async () => {
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
    })().then(completeRelease);
    return task;
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
      if (releasing.has(key)) {
        fail(
          extensionDiagnostic(
            'EXTENSION_DUPLICATE_ID',
            'validate',
            `Extension ${manifest.kind}/${manifest.id} is still being disposed`,
            manifest,
          ),
        );
      }
      if (!manifest.environments.includes(options.environment)) {
        fail(
          extensionDiagnostic(
            'EXTENSION_ENVIRONMENT_UNSUPPORTED',
            'validate',
            `Extension ${manifest.kind}/${manifest.id} does not support host environment ${options.environment}`,
            manifest,
            { environment: options.environment },
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
          diagnostics: publishDiagnostic,
        });
      } catch (cause) {
        controller.abort();
        const primaryDiagnostic = disposed
          ? extensionDiagnostic(
              'EXTENSION_REGISTRY_DISPOSED',
              'validate',
              `Registry disposal cancelled initialization of ${manifest.kind}/${manifest.id}`,
              manifest,
              undefined,
              cause,
            )
          : extensionDiagnostic(
              'EXTENSION_INITIALIZE_FAILED',
              'execute',
              `Failed to initialize extension ${manifest.kind}/${manifest.id}`,
              manifest,
              undefined,
              cause,
            );
        publishDiagnostic(primaryDiagnostic);
        const cleanupDiagnostics = await release(record);
        pending.delete(key);
        completePending(cleanupDiagnostics);
        throw new ExtensionKernelError([primaryDiagnostic, ...cleanupDiagnostics]);
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
      if (query.environment !== options.environment) {
        fail(
          extensionDiagnostic(
            'EXTENSION_ENVIRONMENT_UNSUPPORTED',
            'validate',
            `Registry host ${options.environment} cannot resolve ${query.environment} capabilities`,
            query.id === undefined ? undefined : manifestForError(kind, query.id),
            { environment: query.environment, hostEnvironment: options.environment },
          ),
        );
      }
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
      const initializing = [...pending.values()].sort((left, right) =>
        compareAscii(left.key, right.key),
      );
      const activeRecords = [...records.values()].sort(compareRecords);
      for (const record of activeRecords) unpublish(record);
      for (const registration of initializing) registration.controller.abort();
      for (const record of activeRecords) record.controller.abort();
      const priorReleases = new Map(releasing);
      const cleanupKeys = [
        ...new Set([...releasing.keys(), ...activeRecords.map(({ key }) => key)]),
      ].sort(compareAscii);
      const cleanupKeySet = new Set(cleanupKeys);
      const activeByKey = new Map(activeRecords.map((record) => [record.key, record]));
      disposePromise = (async () => {
        const diagnostics: ExtensionDiagnostic[] = [];
        for (const key of cleanupKeys) {
          const record = activeByKey.get(key);
          const cleanup = record === undefined ? priorReleases.get(key) : release(record);
          if (cleanup !== undefined) diagnostics.push(...(await cleanup));
        }
        for (const registration of initializing) {
          const completionDiagnostics = await registration.completion;
          if (!cleanupKeySet.has(registration.key)) {
            diagnostics.push(...completionDiagnostics);
          }
        }
        disposeComplete = true;
        return Object.freeze(diagnostics);
      })();
      return disposePromise;
    },
  };
}
