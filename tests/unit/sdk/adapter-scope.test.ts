import { describe, expect, it, vi } from 'vitest';
import {
  createAdapterRegistry,
  type AdapterRegistration,
  type IsolatedWorkerTransport,
  type SolverAdapter,
} from '../../../src/sdk/adapters';
import { createCapabilityGrant, type CapabilityGrant } from '../../../src/sdk/trust';

function trustedSolver(
  invoke: SolverAdapter['invoke'],
  capabilities: readonly string[] = ['solve'],
): AdapterRegistration<'solver'> {
  return {
    manifest: {
      id: 'linear',
      apiVersion: '1.0',
      kind: 'solver',
      environments: ['browser'],
      execution: 'trusted-main',
      priority: 0,
      capabilities,
    },
    implementation: { invoke },
  };
}

describe('AdapterScope invocation boundary', () => {
  it('grants only declared capabilities and validates caller-owned result schemas', async () => {
    const direct = vi.fn(async ({ input }: { readonly input: unknown }) => ({
      value: input,
    }));
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct));
    const resolution = registry.resolve('solver');
    const controller = new AbortController();
    const grant = createCapabilityGrant(['solve']);
    const scope = registry.createScope({
      documentId: 'document-1',
      signal: controller.signal,
      grant,
      limits: {
        maxConcurrentInvocations: 1,
        maxDurationMs: 1_000,
        maxInputBytes: 1_024,
        maxOutputBytes: 1_024,
      },
    });

    await expect(
      scope.invoke(resolution, {
        capability: 'inspect',
        input: {},
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'CAPABILITY_DENIED' });
    await expect(
      scope.invoke(resolution, {
        capability: 'solve',
        input: 42,
        validateResult: (value): value is { readonly value: number } =>
          typeof value === 'object' &&
          value !== null &&
          'value' in value &&
          typeof value.value === 'number',
      }),
    ).resolves.toEqual({ value: 42 });
    expect(direct).toHaveBeenCalledWith(
      { capability: 'solve', input: 42 },
      expect.objectContaining({
        documentId: 'document-1',
        signal: expect.any(AbortSignal),
      }),
    );
    expect(Object.isFrozen(grant)).toBe(true);
    expect(Object.isFrozen(grant.capabilities)).toBe(true);
  });

  it('rejects an already-aborted scope before scheduling adapter work', async () => {
    const direct = vi.fn(async () => null);
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct));
    const parent = new AbortController();
    parent.abort(new Error('cancelled before scope creation'));
    const scope = registry.createScope({
      signal: parent.signal,
      grant: createCapabilityGrant(['solve']),
    });

    const rejected = scope.invoke(registry.resolve('solver'), {
      capability: 'solve',
      input: null,
      validateResult: () => true,
    });
    await expect(rejected).rejects.toMatchObject({ code: 'ADAPTER_INVOCATION_ABORTED' });
    expect(direct).not.toHaveBeenCalled();
  });

  it('snapshots hostile invocation fields once before authorization and execution', async () => {
    const direct = vi.fn(async ({ capability, input }) => ({ capability, input }));
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct, ['solve', 'inspect']));
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });
    let capabilityReads = 0;
    let inputReads = 0;
    let validatorReads = 0;
    const invocation = {
      get capability() {
        capabilityReads += 1;
        return capabilityReads === 1 ? 'solve' : 'inspect';
      },
      get input() {
        inputReads += 1;
        return inputReads === 1 ? { value: 'first' } : { value: 'second' };
      },
      get validateResult() {
        validatorReads += 1;
        return () => true;
      },
    };

    await expect(scope.invoke(registry.resolve('solver'), invocation as never)).resolves.toEqual({
      capability: 'solve',
      input: { value: 'first' },
    });
    expect({ capabilityReads, inputReads, validatorReads }).toEqual({
      capabilityReads: 1,
      inputReads: 1,
      validatorReads: 1,
    });
    expect(direct).toHaveBeenCalledWith(
      { capability: 'solve', input: { value: 'first' } },
      expect.any(Object),
    );
  });

  it('snapshots grants instead of trusting a caller-supplied allows method', async () => {
    const direct = vi.fn(async () => null);
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct, ['solve', 'inspect']));
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: {
        capabilities: ['inspect'],
        allows: () => true,
      },
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: null,
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'CAPABILITY_DENIED' });
    expect(direct).not.toHaveBeenCalled();
  });

  it('reads scope options once and uses one parent signal identity through listener cleanup', async () => {
    const firstParent = new AbortController();
    const secondParent = new AbortController();
    const addListener = vi.spyOn(firstParent.signal, 'addEventListener');
    const removeListener = vi.spyOn(firstParent.signal, 'removeEventListener');
    const reads = { signal: 0, documentId: 0, grant: 0, limits: 0 };
    const direct = vi.fn(async (_request, context) => ({
      documentId: context.documentId,
    }));
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct));
    const scope = registry.createScope({
      get signal() {
        reads.signal += 1;
        return reads.signal === 1 ? firstParent.signal : secondParent.signal;
      },
      get documentId() {
        reads.documentId += 1;
        return reads.documentId === 1 ? 'first-document' : 'changed-document';
      },
      get grant() {
        reads.grant += 1;
        return createCapabilityGrant(['solve']);
      },
      get limits() {
        reads.limits += 1;
        return undefined;
      },
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: null,
        validateResult: (value): value is { readonly documentId: string } =>
          typeof value === 'object' &&
          value !== null &&
          'documentId' in value &&
          typeof value.documentId === 'string',
      }),
    ).resolves.toEqual({ documentId: 'first-document' });
    expect(reads).toEqual({ signal: 1, documentId: 1, grant: 1, limits: 1 });
    expect(addListener).toHaveBeenCalledTimes(1);

    await scope.dispose();
    expect(removeListener).toHaveBeenCalledWith('abort', expect.any(Function));
    expect(reads).toEqual({ signal: 1, documentId: 1, grant: 1, limits: 1 });
  });

  it('normalizes hostile scope options and binds cancellation to the first signal snapshot', async () => {
    const firstParent = new AbortController();
    const secondParent = new AbortController();
    let signalReads = 0;
    const direct = vi.fn(async () => null);
    const diagnostics = vi.fn();
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      diagnostics,
    });
    await registry.register(trustedSolver(direct));
    const scope = registry.createScope({
      get signal() {
        signalReads += 1;
        return signalReads === 1 ? firstParent.signal : secondParent.signal;
      },
      grant: createCapabilityGrant(['solve']),
    });
    firstParent.abort(new Error('first parent cancelled'));

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: null,
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_INVOCATION_ABORTED' });
    expect(signalReads).toBe(1);
    expect(direct).not.toHaveBeenCalled();

    expect(() =>
      registry.createScope({
        signal: new AbortController().signal,
        get grant(): CapabilityGrant {
          throw new Error('hostile grant');
        },
      }),
    ).toThrowError(expect.objectContaining({ code: 'ADAPTER_OPTIONS_INVALID' }));
    expect(diagnostics).toHaveBeenCalledWith(
      expect.objectContaining({ code: 'ADAPTER_OPTIONS_INVALID' }),
    );
  });

  it.each([
    new Map([['value', 1]]),
    new Set([1]),
    () => undefined,
    { toJSON: () => ({ safe: true }) },
    Object.defineProperty({}, 'value', { enumerable: true, get: () => 1 }),
  ])('rejects non-plain or executable invocation input %#', async (input) => {
    const direct = vi.fn(async () => null);
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(trustedSolver(direct));
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input,
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_INPUT_INVALID' });
    expect(direct).not.toHaveBeenCalled();
  });

  it('rejects cycles and unsafe results while passing a frozen detached input snapshot', async () => {
    const cycle: { self?: unknown } = {};
    cycle.self = cycle;
    const seen: unknown[] = [];
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(
      trustedSolver(async ({ input }) => {
        seen.push(input);
        return new Map([['unsafe', true]]);
      }),
    );
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: cycle,
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_INPUT_INVALID' });

    const original = { nested: { value: 1 } };
    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: original,
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_RESULT_INVALID' });
    expect(seen[0]).not.toBe(original);
    expect(Object.isFrozen(seen[0])).toBe(true);
    expect(Object.isFrozen((seen[0] as { nested: object }).nested)).toBe(true);
  });

  it('normalizes adapter exceptions and invalid results to invocation diagnostics', async () => {
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(
      trustedSolver(async () => {
        throw new Error('solver exploded');
      }),
    );
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: {},
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({
      code: 'ADAPTER_INVOCATION_FAILED',
      diagnostic: expect.objectContaining({ cause: expect.any(Error) }),
    });

    const invalidRegistry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await invalidRegistry.register(trustedSolver(async () => 'invalid'));
    const invalidScope = invalidRegistry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });
    await expect(
      invalidScope.invoke(invalidRegistry.resolve('solver'), {
        capability: 'solve',
        input: {},
        validateResult: (value): value is number => typeof value === 'number',
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_RESULT_INVALID' });
  });

  it('enforces input, output, concurrency, timeout, and parent abort budgets', async () => {
    vi.useFakeTimers();
    try {
      const registry = createAdapterRegistry({
        apiVersion: '1.0',
        environment: 'browser',
      });
      await registry.register(
        trustedSolver(
          async (_request, context) =>
            new Promise((resolve, reject) => {
              context.signal.addEventListener('abort', () => reject(context.signal.reason), {
                once: true,
              });
            }),
        ),
      );
      const parent = new AbortController();
      const scope = registry.createScope({
        signal: parent.signal,
        grant: createCapabilityGrant(['solve']),
        limits: {
          maxConcurrentInvocations: 1,
          maxDurationMs: 20,
          maxInputBytes: 16,
          maxOutputBytes: 16,
        },
      });
      const resolution = registry.resolve('solver');

      await expect(
        scope.invoke(resolution, {
          capability: 'solve',
          input: 'x'.repeat(32),
          validateResult: () => true,
        }),
      ).rejects.toMatchObject({ code: 'ADAPTER_LIMIT_EXCEEDED' });

      const outputRegistry = createAdapterRegistry({
        apiVersion: '1.0',
        environment: 'browser',
      });
      await outputRegistry.register(trustedSolver(async () => 'x'.repeat(32)));
      const outputScope = outputRegistry.createScope({
        signal: new AbortController().signal,
        grant: createCapabilityGrant(['solve']),
        limits: {
          maxConcurrentInvocations: 1,
          maxDurationMs: 20,
          maxInputBytes: 16,
          maxOutputBytes: 16,
        },
      });
      await expect(
        outputScope.invoke(outputRegistry.resolve('solver'), {
          capability: 'solve',
          input: null,
          validateResult: (value): value is string => typeof value === 'string',
        }),
      ).rejects.toMatchObject({ code: 'ADAPTER_LIMIT_EXCEEDED' });

      const pending = scope.invoke(resolution, {
        capability: 'solve',
        input: null,
        validateResult: () => true,
      });
      const pendingResult = expect(pending).rejects.toMatchObject({
        code: 'ADAPTER_INVOCATION_TIMEOUT',
      });
      await expect(
        scope.invoke(resolution, {
          capability: 'solve',
          input: null,
          validateResult: () => true,
        }),
      ).rejects.toMatchObject({ code: 'ADAPTER_LIMIT_EXCEEDED' });
      await vi.advanceTimersByTimeAsync(20);
      await pendingResult;

      const second = scope.invoke(resolution, {
        capability: 'solve',
        input: null,
        validateResult: () => true,
      });
      const secondResult = expect(second).rejects.toMatchObject({
        code: 'ADAPTER_INVOCATION_ABORTED',
      });
      parent.abort(new Error('host cancelled'));
      await secondResult;
    } finally {
      vi.useRealTimers();
    }
  });

  it('routes isolated-worker declarations only through the injected transport', async () => {
    const direct = vi.fn();
    const transport: IsolatedWorkerTransport = {
      invoke: vi.fn(async (request) => ({ adapterId: request.adapterId })),
      terminate: vi.fn(async () => undefined),
    };
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      isolatedWorkerTransport: transport,
    });
    await registry.register({
      manifest: {
        id: 'isolated',
        apiVersion: '1.0',
        kind: 'solver',
        environments: ['browser'],
        execution: 'isolated-worker',
        priority: 0,
        capabilities: ['solve'],
      },
      descriptor: { workerId: 'solver-worker' },
    });
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });

    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: { objective: 1 },
        validateResult: (value): value is { readonly adapterId: string } =>
          typeof value === 'object' &&
          value !== null &&
          'adapterId' in value &&
          typeof value.adapterId === 'string',
      }),
    ).resolves.toEqual({ adapterId: 'isolated' });
    expect(direct).not.toHaveBeenCalled();
    expect(registry.resolve('solver')).not.toHaveProperty('implementation');
    expect(transport.invoke).toHaveBeenCalledWith(
      expect.objectContaining({
        adapterId: 'isolated',
        workerId: 'solver-worker',
        kind: 'solver',
        capability: 'solve',
        input: { objective: 1 },
      }),
      expect.any(AbortSignal),
    );
  });

  it('rejects forged, foreign, and unregistered stale resolutions', async () => {
    const first = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    const second = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    const unregister = await first.register(trustedSolver(async () => null));
    await second.register(trustedSolver(async () => null));
    const firstScope = first.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });
    const firstResolution = first.resolve('solver');
    const secondResolution = second.resolve('solver');
    const invocation = {
      capability: 'solve',
      input: null,
      validateResult: () => true,
    };

    await expect(
      firstScope.invoke(
        {
          manifest: firstResolution.manifest,
          reason: 'explicit-id',
        } as never,
        invocation,
      ),
    ).rejects.toMatchObject({ code: 'ADAPTER_RESOLUTION_INVALID' });
    await expect(firstScope.invoke(secondResolution, invocation)).rejects.toMatchObject({
      code: 'ADAPTER_RESOLUTION_INVALID',
    });

    await unregister();
    await expect(firstScope.invoke(firstResolution, invocation)).rejects.toMatchObject({
      code: 'ADAPTER_RESOLUTION_INVALID',
    });
  });

  it('aborts active work and removes resources exactly once when disposed', async () => {
    const aborted = vi.fn();
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register(
      trustedSolver(
        async (_request, context) =>
          new Promise((resolve) => {
            context.signal.addEventListener('abort', () => {
              aborted();
              resolve('cancelled');
            });
          }),
      ),
    );
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });
    const pending = scope.invoke(registry.resolve('solver'), {
      capability: 'solve',
      input: {},
      validateResult: (value): value is string => typeof value === 'string',
    });
    const pendingResult = expect(pending).rejects.toMatchObject({
      code: 'ADAPTER_INVOCATION_ABORTED',
    });
    await Promise.resolve();

    const first = scope.dispose();
    const second = scope.dispose();
    await expect(first).resolves.toEqual([]);
    await expect(second).resolves.toEqual([]);
    await pendingResult;
    expect(aborted).toHaveBeenCalledTimes(1);
    await expect(
      scope.invoke(registry.resolve('solver'), {
        capability: 'solve',
        input: {},
        validateResult: () => true,
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_SCOPE_DISPOSED' });
  });

  it('keeps a timed-out trusted invocation in the concurrency ledger until execution settles', async () => {
    vi.useFakeTimers();
    try {
      let settleExecution!: (value: string) => void;
      const execution = new Promise<string>((resolve) => {
        settleExecution = resolve;
      });
      const registry = createAdapterRegistry({
        apiVersion: '1.0',
        environment: 'browser',
      });
      await registry.register(trustedSolver(async () => execution));
      const scope = registry.createScope({
        signal: new AbortController().signal,
        grant: createCapabilityGrant(['solve']),
        limits: {
          maxConcurrentInvocations: 1,
          maxDurationMs: 10,
          maxInputBytes: 1_024,
          maxOutputBytes: 1_024,
        },
      });
      const resolution = registry.resolve('solver');
      const pending = scope.invoke(resolution, {
        capability: 'solve',
        input: null,
        validateResult: (value): value is string => typeof value === 'string',
      });
      const pendingResult = expect(pending).rejects.toMatchObject({
        code: 'ADAPTER_INVOCATION_TIMEOUT',
      });
      await vi.advanceTimersByTimeAsync(10);
      await pendingResult;

      await expect(
        scope.invoke(resolution, {
          capability: 'solve',
          input: null,
          validateResult: () => true,
        }),
      ).rejects.toMatchObject({ code: 'ADAPTER_LIMIT_EXCEEDED' });

      let disposed = false;
      const disposal = scope.dispose().then((diagnostics) => {
        disposed = true;
        return diagnostics;
      });
      await Promise.resolve();
      expect(disposed).toBe(false);

      settleExecution('late result');
      await expect(disposal).resolves.toEqual([]);
    } finally {
      vi.useRealTimers();
    }
  });

  it('waits for isolated worker termination acknowledgement during disposal', async () => {
    let acknowledgeTermination!: () => void;
    const termination = new Promise<void>((resolve) => {
      acknowledgeTermination = resolve;
    });
    const transport: IsolatedWorkerTransport = {
      invoke: vi.fn(() => new Promise(() => undefined)),
      terminate: vi.fn(() => termination),
    };
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      isolatedWorkerTransport: transport,
    });
    await registry.register({
      manifest: {
        id: 'isolated',
        apiVersion: '1.0',
        kind: 'solver',
        environments: ['browser'],
        execution: 'isolated-worker',
        priority: 0,
        capabilities: ['solve'],
      },
      descriptor: { workerId: 'solver-worker' },
    });
    const scope = registry.createScope({
      signal: new AbortController().signal,
      grant: createCapabilityGrant(['solve']),
    });
    const pending = scope.invoke(registry.resolve('solver'), {
      capability: 'solve',
      input: null,
      validateResult: () => true,
    });
    const pendingResult = expect(pending).rejects.toMatchObject({
      code: 'ADAPTER_INVOCATION_ABORTED',
    });
    await Promise.resolve();

    let disposed = false;
    const disposal = scope.dispose().then((diagnostics) => {
      disposed = true;
      return diagnostics;
    });
    await Promise.resolve();
    expect(transport.terminate).toHaveBeenCalledTimes(1);
    expect(transport.terminate).toHaveBeenCalledWith('solver-worker');
    expect(disposed).toBe(false);

    acknowledgeTermination();
    await expect(disposal).resolves.toEqual([]);
    await pendingResult;
  });
});
