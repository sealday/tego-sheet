import type { Diagnostic } from '../../document/diagnostics';
import type { JsonValue } from '../../core/types/json';
import { ExtensionKernelError, type ExtensionManifest } from '../../extensions/kernel/manifest';
import type { KernelContext, KernelRegistration } from '../../extensions/kernel/capabilities';
import { createAdapterRegistryKernel } from '../../extensions/kernel/registry';
import './kernel-bridge';
import { adapterDiagnostic, AdapterSdkError, mapKernelError } from './diagnostics';
import { snapshotAdapterManifest } from './manifest';
import { ADAPTER_KINDS } from './manifest';
import { createAdapterScopeRuntime } from './scope';
import type { AdapterScopeResolution } from './scope';
import { snapshotJsonValue } from './json-safe';
import type {
  AdapterByKind,
  AdapterDiagnostic,
  AdapterKind,
  AdapterManifest,
  AdapterQuery,
  AdapterRegistration,
  AdapterRegistry,
  AdapterRegistryOptions,
  AdapterResolution,
  AdapterResolutionQuery,
  AdapterScope,
  AdapterScopeOptions,
  IsolatedWorkerAdapterDescriptor,
  TrustedMainAdapterRegistration,
} from './types';

interface PublicRecord<K extends AdapterKind = AdapterKind> {
  readonly manifest: AdapterManifest<K>;
  readonly implementation?: AdapterByKind[K];
  readonly descriptor?: IsolatedWorkerAdapterDescriptor;
  readonly unregister: () => Promise<readonly AdapterDiagnostic[]>;
}

function keyOf(kind: string, id: string): string {
  return `${kind}\u0000${id}`;
}

function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function compareManifest(left: AdapterManifest, right: AdapterManifest): number {
  return (
    compareAscii(left.kind, right.kind) ||
    right.priority - left.priority ||
    compareAscii(left.id, right.id)
  );
}

/** Creates a public registry facade backed by the existing F5 registry kernel. */
export function createAdapterRegistry(options: AdapterRegistryOptions): AdapterRegistry {
  let configuration: AdapterRegistryOptions;
  try {
    const defaultsSource = options.defaults;
    const defaults: Partial<Record<AdapterKind, string>> = {};
    if (defaultsSource !== undefined) {
      for (const kind of ADAPTER_KINDS) {
        const id = defaultsSource[kind];
        if (id !== undefined) {
          if (typeof id !== 'string' || id.length === 0) {
            throw new TypeError(`Default adapter ${kind} must be a non-empty ID`);
          }
          defaults[kind] = id;
        }
      }
    }
    configuration = Object.freeze({
      apiVersion: options.apiVersion,
      environment: options.environment,
      defaults: Object.freeze(defaults),
      ...(options.diagnostics === undefined ? {} : { diagnostics: options.diagnostics }),
      ...(options.isolatedWorkerTransport === undefined
        ? {}
        : { isolatedWorkerTransport: options.isolatedWorkerTransport }),
    });
  } catch (cause) {
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_OPTIONS_INVALID',
        'validate',
        'Adapter registry options could not be snapshotted',
        { cause },
      ),
    ]);
  }
  const kernel = createAdapterRegistryKernel({
    apiVersion: configuration.apiVersion,
    environment: configuration.environment,
  });
  const records = new Map<string, PublicRecord>();
  const resolutionRecords = new WeakMap<object, PublicRecord>();
  const scopes = new Set<AdapterScope>();
  let disposed = false;
  let disposePromise: Promise<readonly AdapterDiagnostic[]> | undefined;

  const publish = (diagnostic: AdapterDiagnostic): void => {
    try {
      configuration.diagnostics?.(diagnostic);
    } catch {
      // Diagnostic observers cannot own registry control flow.
    }
  };

  const fail = (
    code: AdapterDiagnostic['code'],
    message: string,
    details?: AdapterDiagnostic['details'],
  ): never => {
    const diagnostic = adapterDiagnostic(
      code,
      'resolve',
      message,
      details === undefined ? {} : { details: details as never },
    );
    publish(diagnostic);
    throw new AdapterSdkError([diagnostic]);
  };

  const publicApi: AdapterRegistry = {
    async register<K extends AdapterKind>(
      registration: AdapterRegistration<K>,
    ): Promise<() => Promise<readonly AdapterDiagnostic[]>> {
      if (disposed) {
        fail('ADAPTER_REGISTRY_DISPOSED', 'Cannot register an adapter after disposal');
      }
      let manifest: AdapterManifest<K>;
      let implementation: AdapterByKind[K] | undefined;
      let descriptor: IsolatedWorkerAdapterDescriptor | undefined;
      let initialize: TrustedMainAdapterRegistration<K>['initialize'];
      let dispose: TrustedMainAdapterRegistration<K>['dispose'];
      try {
        if (registration === null || typeof registration !== 'object') {
          throw new TypeError('Adapter registration must be an object');
        }
        const manifestProperty = Object.getOwnPropertyDescriptor(registration, 'manifest');
        if (manifestProperty === undefined || !('value' in manifestProperty)) {
          throw new TypeError('Adapter manifest must be a plain data property');
        }
        manifest = snapshotAdapterManifest(manifestProperty.value as AdapterManifest<K>);
        if (manifest.execution === 'isolated-worker') {
          if (
            Object.hasOwn(registration, 'implementation') ||
            Object.hasOwn(registration, 'initialize') ||
            Object.hasOwn(registration, 'dispose')
          ) {
            throw new TypeError('Isolated adapters cannot register main-thread code');
          }
          const descriptorProperty = Object.getOwnPropertyDescriptor(registration, 'descriptor');
          if (descriptorProperty === undefined || !('value' in descriptorProperty)) {
            throw new TypeError('Isolated adapters require a transport-owned descriptor');
          }
          const descriptorSnapshot = snapshotJsonValue(
            descriptorProperty.value,
            'adapter.descriptor',
          );
          if (
            descriptorSnapshot === null ||
            Array.isArray(descriptorSnapshot) ||
            typeof descriptorSnapshot !== 'object'
          ) {
            throw new TypeError('Isolated adapter descriptor must contain one workerId');
          }
          const descriptorObject = descriptorSnapshot as Readonly<Record<string, JsonValue>>;
          const workerId = descriptorObject.workerId;
          if (
            typeof workerId !== 'string' ||
            workerId.length === 0 ||
            Object.keys(descriptorObject).length !== 1
          ) {
            throw new TypeError('Isolated adapter descriptor must contain one workerId');
          }
          descriptor = descriptorObject as unknown as IsolatedWorkerAdapterDescriptor;
        } else {
          if (Object.hasOwn(registration, 'descriptor')) {
            throw new TypeError('Trusted adapters cannot register a worker descriptor');
          }
          const implementationProperty = Object.getOwnPropertyDescriptor(
            registration,
            'implementation',
          );
          if (implementationProperty === undefined || !('value' in implementationProperty)) {
            throw new TypeError('Trusted adapters require a plain implementation property');
          }
          implementation = implementationProperty.value as AdapterByKind[K];
          const initializeProperty = Object.getOwnPropertyDescriptor(registration, 'initialize');
          const disposeProperty = Object.getOwnPropertyDescriptor(registration, 'dispose');
          if (initializeProperty !== undefined && !('value' in initializeProperty)) {
            throw new TypeError('Adapter initialize must be a plain data property');
          }
          if (disposeProperty !== undefined && !('value' in disposeProperty)) {
            throw new TypeError('Adapter dispose must be a plain data property');
          }
          initialize = initializeProperty?.value as typeof initialize;
          dispose = disposeProperty?.value as typeof dispose;
        }
      } catch (error) {
        if (error instanceof AdapterSdkError) {
          error.diagnostics.forEach(publish);
          throw error;
        }
        const normalized = new AdapterSdkError([
          adapterDiagnostic(
            'ADAPTER_MANIFEST_INVALID',
            'validate',
            'Adapter registration is invalid',
            { cause: error },
          ),
        ]);
        normalized.diagnostics.forEach(publish);
        throw normalized;
      }

      let kernelUnregister: () => Promise<readonly Diagnostic[]>;
      try {
        const kernelRegistration = {
          manifest: {
            id: manifest.id,
            apiVersion: manifest.apiVersion,
            kind: manifest.kind,
            environments: manifest.environments,
          } as ExtensionManifest & { readonly kind: K },
          implementation:
            manifest.execution === 'isolated-worker'
              ? Object.freeze({
                  adapterId: manifest.id,
                  workerId: descriptor!.workerId,
                })
              : implementation,
          ...(initialize === undefined
            ? {}
            : {
                initialize: ({ environment, signal }: KernelContext) =>
                  initialize?.({
                    environment,
                    signal,
                    diagnostics: (diagnostic) => {
                      if (
                        diagnostic.domain === 'extension' &&
                        diagnostic.code.startsWith('ADAPTER_')
                      ) {
                        publish(diagnostic as AdapterDiagnostic);
                      }
                    },
                  }),
              }),
          ...(dispose === undefined ? {} : { dispose }),
        } as unknown as KernelRegistration<K>;
        kernelUnregister = await kernel.register(kernelRegistration);
      } catch (error) {
        if (error instanceof ExtensionKernelError) {
          const mapped = mapKernelError(error);
          mapped.diagnostics.forEach(publish);
          throw mapped;
        }
        throw error;
      }

      const key = keyOf(manifest.kind, manifest.id);
      let releasePromise: Promise<readonly AdapterDiagnostic[]> | undefined;
      const unregister = (): Promise<readonly AdapterDiagnostic[]> => {
        if (releasePromise !== undefined) return releasePromise;
        records.delete(key);
        releasePromise = kernelUnregister().then((diagnostics) => {
          const mapped = diagnostics.map((diagnostic) =>
            adapterDiagnostic('ADAPTER_DISPOSE_FAILED', diagnostic.stage, diagnostic.message, {
              manifest: { id: manifest.id },
              ...(diagnostic.cause === undefined ? {} : { cause: diagnostic.cause }),
            }),
          );
          mapped.forEach(publish);
          return Object.freeze(mapped);
        });
        return releasePromise;
      };
      records.set(key, {
        manifest,
        ...(implementation === undefined ? {} : { implementation }),
        ...(descriptor === undefined ? {} : { descriptor }),
        unregister,
      } as PublicRecord);
      return unregister;
    },

    list(query: AdapterQuery = {}): readonly AdapterManifest[] {
      return Object.freeze(
        [...records.values()]
          .map(({ manifest }) => manifest)
          .filter(
            (manifest) =>
              (query.kind === undefined || manifest.kind === query.kind) &&
              (query.capability === undefined ||
                manifest.capabilities.includes(query.capability)) &&
              (query.format === undefined || manifest.formats?.includes(query.format) === true),
          )
          .sort(compareManifest),
      );
    },

    resolve<K extends AdapterKind>(
      kind: K,
      query: AdapterResolutionQuery = {},
    ): AdapterResolution<K> {
      if (disposed) {
        fail('ADAPTER_REGISTRY_DISPOSED', 'Cannot resolve an adapter after disposal');
      }
      const candidates = publicApi
        .list({ kind, capability: query.capability, format: query.format })
        .filter((manifest): manifest is AdapterManifest<K> => manifest.kind === kind);
      let selected: AdapterManifest<K> | undefined;
      let reason: AdapterResolution<K>['reason'] = 'single-match';
      if (query.id !== undefined) {
        selected = candidates.find(({ id }) => id === query.id);
        reason = 'explicit-id';
      } else {
        const configured = configuration.defaults?.[kind];
        if (configured !== undefined) {
          selected = candidates.find(({ id }) => id === configured);
          reason = 'configured-default';
        } else if (candidates.length === 1) {
          selected = candidates[0];
          reason = 'single-match';
        } else if (candidates.length > 1) {
          fail('ADAPTER_AMBIGUOUS', `More than one ${kind} adapter matches; specify an ID`, {
            candidates: candidates.map(({ id }) => id).sort(),
            kind,
          });
        } else {
          reason = 'single-match';
        }
      }
      const chosen: AdapterManifest<K> =
        selected ??
        fail('ADAPTER_NOT_FOUND', `No ${kind} adapter matches the resolution query`, {
          kind,
          ...(query.id === undefined ? {} : { id: query.id }),
        });
      try {
        kernel.resolve(kind, {
          id: chosen.id,
          environment: configuration.environment,
        });
        const record =
          records.get(keyOf(chosen.kind, chosen.id)) ??
          fail('ADAPTER_NOT_FOUND', `Adapter ${chosen.id} is no longer registered`);
        const resolution = Object.freeze({ manifest: chosen, reason });
        resolutionRecords.set(resolution, record);
        return resolution;
      } catch (error) {
        if (error instanceof ExtensionKernelError) {
          const mapped = mapKernelError(error);
          mapped.diagnostics.forEach(publish);
          throw mapped;
        }
        throw error;
      }
    },

    createScope(scopeOptions: AdapterScopeOptions): AdapterScope {
      if (disposed) {
        fail('ADAPTER_REGISTRY_DISPOSED', 'Cannot create an adapter scope after disposal');
      }
      const scope = createAdapterScopeRuntime({
        options: scopeOptions,
        ...(configuration.isolatedWorkerTransport === undefined
          ? {}
          : { transport: configuration.isolatedWorkerTransport }),
        publish,
        onDispose: () => scopes.delete(scope),
        lookupResolution: (resolution): AdapterScopeResolution | undefined => {
          const record = resolutionRecords.get(resolution);
          if (
            record === undefined ||
            records.get(keyOf(record.manifest.kind, record.manifest.id)) !== record
          ) {
            return undefined;
          }
          return record;
        },
      });
      scopes.add(scope);
      return scope;
    },

    dispose(): Promise<readonly AdapterDiagnostic[]> {
      if (disposePromise !== undefined) return disposePromise;
      disposed = true;
      disposePromise = (async () => {
        const diagnostics: AdapterDiagnostic[] = [];
        for (const scope of scopes) diagnostics.push(...(await scope.dispose()));
        const registrations = [...records.values()].reverse();
        for (const record of registrations) diagnostics.push(...(await record.unregister()));
        const kernelDiagnostics = await kernel.dispose();
        for (const diagnostic of kernelDiagnostics) {
          const mapped = adapterDiagnostic(
            'ADAPTER_DISPOSE_FAILED',
            diagnostic.stage,
            diagnostic.message,
            {
              ...(diagnostic.location?.adapterId === undefined
                ? {}
                : { manifest: { id: diagnostic.location.adapterId } }),
              ...(diagnostic.cause === undefined ? {} : { cause: diagnostic.cause }),
            },
          );
          diagnostics.push(mapped);
          publish(mapped);
        }
        return Object.freeze(diagnostics);
      })();
      return disposePromise;
    },
  };

  return publicApi;
}
