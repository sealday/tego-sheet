import { describe, expect, it, vi } from 'vitest';
import {
  createAdapterRegistry,
  type AdapterRegistration,
  type IsolatedWorkerTransport,
  type SolverAdapter,
} from '../../../src/sdk/adapters';
import { createCapabilityGrant } from '../../../src/sdk/trust';

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
      implementation: { invoke: direct },
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
    expect(registry.resolve('solver').implementation).toEqual({
      execution: 'isolated-worker',
      adapterId: 'isolated',
    });
    expect(transport.invoke).toHaveBeenCalledWith(
      expect.objectContaining({
        adapterId: 'isolated',
        kind: 'solver',
        capability: 'solve',
        input: { objective: 1 },
      }),
      expect.any(AbortSignal),
    );
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
});
