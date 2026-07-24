import type { JsonValue } from '../../core/types/json';
import { DEFAULT_ADAPTER_SCOPE_LIMITS, type AdapterScopeLimits } from '../trust';
import { adapterDiagnostic, AdapterSdkError } from './diagnostics';
import type {
  AdapterDiagnostic,
  AdapterKind,
  AdapterResolution,
  AdapterScope,
  AdapterScopeOptions,
  CallableAdapter,
  IsolatedWorkerTransport,
  ScopedAdapterInvocation,
} from './types';

const encoder = new TextEncoder();
const timeoutMarker = Object.freeze({ type: 'adapter-timeout' });
const disposedMarker = Object.freeze({ type: 'adapter-scope-disposed' });

interface AdapterScopeRuntimeOptions {
  readonly options: AdapterScopeOptions;
  readonly transport?: IsolatedWorkerTransport;
  readonly publish: (diagnostic: AdapterDiagnostic) => void;
  readonly onDispose: () => void;
}

function limitsSnapshot(overrides: Partial<AdapterScopeLimits> | undefined): AdapterScopeLimits {
  const limits = Object.freeze({ ...DEFAULT_ADAPTER_SCOPE_LIMITS, ...overrides });
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
    throw new TypeError('Adapter scope limits must be positive finite integers');
  }
  return limits;
}

function encodedSize(value: unknown): number {
  const serialized = JSON.stringify(value);
  return serialized === undefined ? 0 : encoder.encode(serialized).byteLength;
}

function callable(value: unknown): value is CallableAdapter {
  return (
    value !== null &&
    typeof value === 'object' &&
    typeof (value as Partial<CallableAdapter>).invoke === 'function'
  );
}

export function createAdapterScopeRuntime(runtime: AdapterScopeRuntimeOptions): AdapterScope {
  const limits = limitsSnapshot(runtime.options.limits);
  const controller = new AbortController();
  const active = new Set<Promise<void>>();
  let disposed = false;
  let disposePromise: Promise<readonly AdapterDiagnostic[]> | undefined;

  const abortFromParent = (): void => controller.abort(runtime.options.signal.reason);
  if (runtime.options.signal.aborted) abortFromParent();
  else runtime.options.signal.addEventListener('abort', abortFromParent, { once: true });

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
      if (
        !runtime.options.grant.allows(invocation.capability) ||
        !resolution.manifest.capabilities.includes(invocation.capability)
      ) {
        fail('CAPABILITY_DENIED', `Capability ${invocation.capability} was not granted`, {
          resolution,
          details: { capability: invocation.capability },
        });
      }
      let inputBytes = 0;
      try {
        inputBytes = encodedSize(invocation.input);
      } catch (cause) {
        return fail('ADAPTER_LIMIT_EXCEEDED', 'Adapter input is not JSON serializable', {
          resolution,
          cause,
        });
      }
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
        if (resolution.manifest.execution === 'isolated-worker') {
          if (runtime.transport === undefined) {
            throw new Error('No isolated-worker transport is configured');
          }
          return runtime.transport.invoke(
            Object.freeze({
              adapterId: resolution.manifest.id,
              kind: resolution.manifest.kind,
              capability: invocation.capability,
              input: structuredClone(invocation.input),
              ...(runtime.options.documentId === undefined
                ? {}
                : { documentId: runtime.options.documentId }),
            }),
            invocationController.signal,
          );
        }
        if (!callable(resolution.implementation)) {
          throw new TypeError(`Adapter ${resolution.manifest.id} is not generically callable`);
        }
        return resolution.implementation.invoke(
          Object.freeze({ capability: invocation.capability, input: invocation.input }),
          Object.freeze({
            ...(runtime.options.documentId === undefined
              ? {}
              : { documentId: runtime.options.documentId }),
            signal: invocationController.signal,
          }),
        );
      });
      const settled = Promise.race([execution, abortPromise]);
      const ledger = settled.then(
        () => undefined,
        () => undefined,
      );
      active.add(ledger);

      try {
        const result = await settled;
        let valid = false;
        try {
          valid = invocation.validateResult(result);
        } catch {
          valid = false;
        }
        if (!valid) {
          fail('ADAPTER_RESULT_INVALID', 'Adapter result failed its host-owned schema', {
            resolution,
            details: { capability: invocation.capability },
          });
        }
        let outputBytes = 0;
        try {
          outputBytes = encodedSize(result);
        } catch (cause) {
          return fail('ADAPTER_RESULT_INVALID', 'Adapter result is not JSON serializable', {
            resolution,
            cause,
          });
        }
        if (outputBytes > limits.maxOutputBytes) {
          fail('ADAPTER_LIMIT_EXCEEDED', 'Adapter result exceeds the scope byte limit', {
            resolution,
            details: { actual: outputBytes, limit: limits.maxOutputBytes },
          });
        }
        return result as Result;
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
          details: { capability: invocation.capability },
        });
      } finally {
        clearTimeout(timeout);
        controller.signal.removeEventListener('abort', abortInvocation);
        active.delete(ledger);
      }
    },

    dispose(): Promise<readonly AdapterDiagnostic[]> {
      if (disposePromise !== undefined) return disposePromise;
      disposed = true;
      runtime.options.signal.removeEventListener('abort', abortFromParent);
      controller.abort(disposedMarker);
      disposePromise = Promise.all(active).then(() => {
        runtime.onDispose();
        return Object.freeze([]);
      });
      return disposePromise;
    },
  };
}
