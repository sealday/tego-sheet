import type { JsonValue } from '../../core/types/json';
import {
  createCapabilityGrant,
  DEFAULT_ADAPTER_SCOPE_LIMITS,
  type AdapterScopeLimits,
  type CapabilityGrant,
} from '../trust';
import { adapterDiagnostic, AdapterSdkError } from './diagnostics';
import { jsonSnapshotBytes, snapshotJsonValue } from './json-safe';
import type {
  AdapterDiagnostic,
  AdapterByKind,
  AdapterKind,
  AdapterManifest,
  AdapterResolution,
  AdapterScope,
  AdapterScopeOptions,
  CallableAdapter,
  IsolatedWorkerTransport,
  IsolatedWorkerAdapterDescriptor,
  ScopedAdapterInvocation,
} from './types';

const timeoutMarker = Object.freeze({ type: 'adapter-timeout' });
const disposedMarker = Object.freeze({ type: 'adapter-scope-disposed' });

interface AdapterScopeRuntimeOptions {
  readonly options: AdapterScopeOptions;
  readonly transport?: IsolatedWorkerTransport;
  readonly publish: (diagnostic: AdapterDiagnostic) => void;
  readonly onDispose: () => void;
  readonly lookupResolution: (resolution: AdapterResolution) => AdapterScopeResolution | undefined;
}

/** Registry-private live binding resolved from one opaque public handle. */
export interface AdapterScopeResolution {
  readonly manifest: AdapterManifest;
  readonly implementation?: AdapterByKind[AdapterKind];
  readonly descriptor?: IsolatedWorkerAdapterDescriptor;
}

interface ActiveInvocation {
  readonly adapterId: string;
  readonly execution: AdapterManifest['execution'];
  readonly workerId?: string;
  readonly completion: Promise<void>;
}

interface ScopeConfiguration {
  readonly documentId?: string;
  readonly signal: AbortSignal;
  readonly grant: CapabilityGrant;
  readonly limits: AdapterScopeLimits;
}

function limitsSnapshot(overrides: Partial<AdapterScopeLimits> | undefined): AdapterScopeLimits {
  let limits: AdapterScopeLimits;
  try {
    limits = Object.freeze({
      maxConcurrentInvocations:
        overrides?.maxConcurrentInvocations ??
        DEFAULT_ADAPTER_SCOPE_LIMITS.maxConcurrentInvocations,
      maxDurationMs: overrides?.maxDurationMs ?? DEFAULT_ADAPTER_SCOPE_LIMITS.maxDurationMs,
      maxInputBytes: overrides?.maxInputBytes ?? DEFAULT_ADAPTER_SCOPE_LIMITS.maxInputBytes,
      maxOutputBytes: overrides?.maxOutputBytes ?? DEFAULT_ADAPTER_SCOPE_LIMITS.maxOutputBytes,
    });
  } catch (cause) {
    throw new AdapterSdkError([
      adapterDiagnostic('ADAPTER_OPTIONS_INVALID', 'validate', 'Scope limits could not be read', {
        cause,
      }),
    ]);
  }
  if (
    !Number.isSafeInteger(limits.maxConcurrentInvocations) ||
    limits.maxConcurrentInvocations <= 0 ||
    !Number.isFinite(limits.maxDurationMs) ||
    limits.maxDurationMs <= 0 ||
    !Number.isSafeInteger(limits.maxInputBytes) ||
    limits.maxInputBytes <= 0 ||
    !Number.isSafeInteger(limits.maxOutputBytes) ||
    limits.maxOutputBytes <= 0
  ) {
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_OPTIONS_INVALID',
        'validate',
        'Adapter scope limits must be positive finite integers',
      ),
    ]);
  }
  return limits;
}

function grantSnapshot(grant: CapabilityGrant): CapabilityGrant {
  try {
    return createCapabilityGrant(grant.capabilities);
  } catch (cause) {
    throw new AdapterSdkError([
      adapterDiagnostic('ADAPTER_OPTIONS_INVALID', 'validate', 'Capability grant is invalid', {
        cause,
      }),
    ]);
  }
}

function optionsSnapshot(options: AdapterScopeOptions): Readonly<ScopeConfiguration> {
  try {
    if (options === null || typeof options !== 'object') {
      throw new TypeError('Adapter scope options must be an object');
    }
    const documentId = options.documentId;
    const signal = options.signal;
    const grant = options.grant;
    const limits = options.limits;
    if (documentId !== undefined && typeof documentId !== 'string') {
      throw new TypeError('Adapter scope documentId must be a string');
    }
    if (!(signal instanceof AbortSignal)) {
      throw new TypeError('Adapter scope signal must be an AbortSignal');
    }
    return Object.freeze({
      ...(documentId === undefined ? {} : { documentId }),
      signal,
      grant: grantSnapshot(grant),
      limits: limitsSnapshot(limits),
    });
  } catch (cause) {
    if (cause instanceof AdapterSdkError && cause.code === 'ADAPTER_OPTIONS_INVALID') {
      throw cause;
    }
    throw new AdapterSdkError([
      adapterDiagnostic(
        'ADAPTER_OPTIONS_INVALID',
        'validate',
        'Adapter scope options could not be snapshotted',
        { cause },
      ),
    ]);
  }
}

function callable(value: unknown): value is CallableAdapter {
  return (
    value !== null &&
    typeof value === 'object' &&
    typeof (value as Partial<CallableAdapter>).invoke === 'function'
  );
}

export function createAdapterScopeRuntime(runtime: AdapterScopeRuntimeOptions): AdapterScope {
  const configuration = optionsSnapshot(runtime.options);
  const { grant, limits } = configuration;
  const controller = new AbortController();
  const active = new Set<ActiveInvocation>();
  let disposed = false;
  let disposePromise: Promise<readonly AdapterDiagnostic[]> | undefined;
  let parentListenerAttached = false;

  const abortFromParent = (): void => controller.abort(configuration.signal.reason);
  if (configuration.signal.aborted) abortFromParent();
  else {
    configuration.signal.addEventListener('abort', abortFromParent, { once: true });
    parentListenerAttached = true;
  }

  const fail = (
    code: AdapterDiagnostic['code'],
    message: string,
    options: {
      readonly resolution?: AdapterResolution;
      readonly cause?: unknown;
      readonly details?: Readonly<Record<string, JsonValue>>;
    } = {},
  ): never => {
    const diagnostic = adapterDiagnostic(code, 'execute', message, {
      ...(options.resolution === undefined
        ? {}
        : { manifest: { id: options.resolution.manifest.id } }),
      ...(options.cause === undefined ? {} : { cause: options.cause }),
      ...(options.details === undefined ? {} : { details: options.details }),
    });
    runtime.publish(diagnostic);
    throw new AdapterSdkError([diagnostic]);
  };

  return {
    async invoke<K extends AdapterKind, Result>(
      resolution: AdapterResolution<K>,
      invocation: ScopedAdapterInvocation<Result>,
    ): Promise<Result> {
      if (disposed) {
        fail('ADAPTER_SCOPE_DISPOSED', 'Adapter scope has been disposed', { resolution });
      }
      if (controller.signal.aborted) {
        fail('ADAPTER_INVOCATION_ABORTED', 'Adapter invocation was aborted', {
          resolution,
          cause: controller.signal.reason,
        });
      }
      const binding =
        runtime.lookupResolution(resolution) ??
        fail(
          'ADAPTER_RESOLUTION_INVALID',
          'Adapter resolution is forged, foreign, or no longer registered',
        );
      let capability: string;
      let input: unknown;
      let validateResult: ScopedAdapterInvocation<Result>['validateResult'];
      try {
        capability = invocation.capability;
        input = invocation.input;
        validateResult = invocation.validateResult;
      } catch (cause) {
        return fail('ADAPTER_INPUT_INVALID', 'Adapter invocation could not be read safely', {
          resolution,
          cause,
        });
      }
      if (typeof capability !== 'string' || typeof validateResult !== 'function') {
        fail('ADAPTER_INPUT_INVALID', 'Adapter invocation has invalid capability or validator', {
          resolution,
        });
      }
      if (!grant.allows(capability) || !binding.manifest.capabilities.includes(capability)) {
        fail('CAPABILITY_DENIED', `Capability ${capability} was not granted`, {
          resolution,
          details: { capability },
        });
      }
      let inputSnapshot: JsonValue;
      try {
        inputSnapshot = snapshotJsonValue(input, 'adapter.input');
      } catch (cause) {
        return fail('ADAPTER_INPUT_INVALID', 'Adapter input is not strict JSON data', {
          resolution,
          cause,
        });
      }
      const inputBytes = jsonSnapshotBytes(inputSnapshot);
      if (inputBytes > limits.maxInputBytes) {
        fail('ADAPTER_LIMIT_EXCEEDED', 'Adapter input exceeds the scope byte limit', {
          resolution,
          details: { actual: inputBytes, limit: limits.maxInputBytes },
        });
      }
      if (active.size >= limits.maxConcurrentInvocations) {
        fail('ADAPTER_LIMIT_EXCEEDED', 'Adapter scope concurrency limit exceeded', {
          resolution,
          details: { limit: limits.maxConcurrentInvocations },
        });
      }

      const invocationController = new AbortController();
      const abortInvocation = (): void => invocationController.abort(controller.signal.reason);
      if (controller.signal.aborted) abortInvocation();
      else controller.signal.addEventListener('abort', abortInvocation, { once: true });
      const timeout = setTimeout(
        () => invocationController.abort(timeoutMarker),
        limits.maxDurationMs,
      );
      const abortPromise = new Promise<never>((_resolve, reject) => {
        invocationController.signal.addEventListener(
          'abort',
          () => reject(invocationController.signal.reason),
          { once: true },
        );
      });

      const execution = Promise.resolve().then(async () => {
        if (binding.manifest.execution === 'isolated-worker') {
          if (runtime.transport === undefined || binding.descriptor === undefined) {
            throw new Error('No isolated-worker transport is configured');
          }
          return runtime.transport.invoke(
            Object.freeze({
              adapterId: binding.manifest.id,
              workerId: binding.descriptor.workerId,
              kind: binding.manifest.kind,
              capability,
              input: inputSnapshot,
              ...(configuration.documentId === undefined
                ? {}
                : { documentId: configuration.documentId }),
            }),
            invocationController.signal,
          );
        }
        if (!callable(binding.implementation)) {
          throw new TypeError(`Adapter ${binding.manifest.id} is not generically callable`);
        }
        return binding.implementation.invoke(
          Object.freeze({ capability, input: inputSnapshot }),
          Object.freeze({
            ...(configuration.documentId === undefined
              ? {}
              : { documentId: configuration.documentId }),
            signal: invocationController.signal,
          }),
        );
      });
      const settled = Promise.race([execution, abortPromise]);
      const completion = execution.then(
        () => undefined,
        () => undefined,
      );
      const ledger = Object.freeze({
        adapterId: binding.manifest.id,
        execution: binding.manifest.execution,
        ...(binding.descriptor === undefined ? {} : { workerId: binding.descriptor.workerId }),
        completion,
      });
      active.add(ledger);
      void completion.finally(() => {
        active.delete(ledger);
      });

      try {
        const result = await settled;
        let resultSnapshot: JsonValue;
        try {
          resultSnapshot = snapshotJsonValue(result, 'adapter.result');
        } catch (cause) {
          return fail('ADAPTER_RESULT_INVALID', 'Adapter result is not strict JSON data', {
            resolution,
            cause,
          });
        }
        let valid = false;
        try {
          valid = validateResult(resultSnapshot);
        } catch {
          valid = false;
        }
        if (!valid) {
          fail('ADAPTER_RESULT_INVALID', 'Adapter result failed its host-owned schema', {
            resolution,
            details: { capability },
          });
        }
        const outputBytes = jsonSnapshotBytes(resultSnapshot);
        if (outputBytes > limits.maxOutputBytes) {
          fail('ADAPTER_LIMIT_EXCEEDED', 'Adapter result exceeds the scope byte limit', {
            resolution,
            details: { actual: outputBytes, limit: limits.maxOutputBytes },
          });
        }
        return resultSnapshot as Result;
      } catch (cause) {
        if (cause instanceof AdapterSdkError) throw cause;
        if (invocationController.signal.aborted) {
          const reason = invocationController.signal.reason;
          fail(
            reason === timeoutMarker ? 'ADAPTER_INVOCATION_TIMEOUT' : 'ADAPTER_INVOCATION_ABORTED',
            reason === timeoutMarker
              ? 'Adapter invocation exceeded its duration limit'
              : 'Adapter invocation was aborted',
            { resolution, cause: reason },
          );
        }
        return fail('ADAPTER_INVOCATION_FAILED', 'Adapter invocation failed', {
          resolution,
          cause,
          details: { capability },
        });
      } finally {
        clearTimeout(timeout);
        controller.signal.removeEventListener('abort', abortInvocation);
      }
    },

    dispose(): Promise<readonly AdapterDiagnostic[]> {
      if (disposePromise !== undefined) return disposePromise;
      disposed = true;
      if (parentListenerAttached) {
        configuration.signal.removeEventListener('abort', abortFromParent);
        parentListenerAttached = false;
      }
      controller.abort(disposedMarker);
      disposePromise = (async () => {
        const diagnostics: AdapterDiagnostic[] = [];
        const terminations = new Map<string, Promise<void>>();
        const waitForInvocation = (invocation: ActiveInvocation): Promise<void> => {
          if (
            invocation.execution !== 'isolated-worker' ||
            invocation.workerId === undefined ||
            runtime.transport === undefined
          ) {
            return invocation.completion;
          }
          let termination = terminations.get(invocation.workerId);
          if (termination === undefined) {
            const workerId = invocation.workerId;
            termination = Promise.resolve()
              .then(() => runtime.transport!.terminate(workerId))
              .catch((cause) => {
                const diagnostic = adapterDiagnostic(
                  'ADAPTER_DISPOSE_FAILED',
                  'dispose',
                  `Isolated worker ${workerId} could not be terminated`,
                  {
                    manifest: { id: invocation.adapterId },
                    cause,
                  },
                );
                diagnostics.push(diagnostic);
                runtime.publish(diagnostic);
              });
            terminations.set(workerId, termination);
          }
          return Promise.race([invocation.completion, termination]).finally(() => {
            active.delete(invocation);
          });
        };
        await Promise.all([...active].map(waitForInvocation));
        runtime.onDispose();
        return Object.freeze(diagnostics);
      })();
      return disposePromise;
    },
  };
}
