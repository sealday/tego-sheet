import { describe, expect, it, vi } from 'vitest';
import {
  ADAPTER_KINDS,
  AdapterSdkError,
  createAdapterRegistry,
  type AdapterRegistration,
  type SolverAdapter,
} from '../../../src/sdk/adapters';

function solver(
  id: string,
  options: Partial<AdapterRegistration<'solver'>> = {},
): AdapterRegistration<'solver'> {
  return {
    manifest: {
      id,
      apiVersion: '1.0',
      kind: 'solver',
      environments: ['browser'],
      execution: 'trusted-main',
      priority: 0,
      capabilities: ['solve'],
    },
    implementation: {
      invoke: async ({ input }) => input,
    } satisfies SolverAdapter,
    ...options,
  };
}

describe('public AdapterRegistry facade', () => {
  it('publishes the bounded core and host-integration adapter kind entries', () => {
    expect(ADAPTER_KINDS).toEqual(
      expect.arrayContaining([
        'workbook-reader',
        'workbook-writer',
        'resource-resolver',
        'output',
        'solver',
        'persistence',
        'collaboration',
        'permission',
        'comments',
        'version-history',
        'ai-command',
      ]),
    );
  });

  it('publishes frozen public manifests while resolving through the F5 kernel', async () => {
    const registry = createAdapterRegistry({
      apiVersion: '1.2',
      environment: 'browser',
    });
    const registration = solver('linear', {
      manifest: {
        ...solver('linear').manifest,
        apiVersion: '1.1',
        priority: 10,
        formats: ['linear'],
      },
    });

    await registry.register(registration);

    const listed = registry.list({ kind: 'solver' });
    expect(listed).toEqual([registration.manifest]);
    expect(Object.isFrozen(listed)).toBe(true);
    expect(Object.isFrozen(listed[0])).toBe(true);
    expect(Object.isFrozen(listed[0]!.capabilities)).toBe(true);
    expect(Object.isFrozen(listed[0]!.formats)).toBe(true);
    expect(registry.resolve('solver', { id: 'linear' })).toMatchObject({
      manifest: registration.manifest,
      implementation: registration.implementation,
      reason: 'explicit-id',
    });
  });

  it('maps kernel duplicate, version, and environment failures to stable public diagnostics', async () => {
    const diagnostics = vi.fn();
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      diagnostics,
    });
    await registry.register(solver('linear'));

    await expect(registry.register(solver('linear'))).rejects.toMatchObject({
      code: 'ADAPTER_DUPLICATE_ID',
    });
    await expect(
      registry.register(
        solver('future', {
          manifest: { ...solver('future').manifest, apiVersion: '2.0' },
        }),
      ),
    ).rejects.toMatchObject({ code: 'ADAPTER_VERSION_UNSUPPORTED' });
    await expect(
      registry.register(
        solver('worker-only', {
          manifest: { ...solver('worker-only').manifest, environments: ['worker'] },
        }),
      ),
    ).rejects.toMatchObject({ code: 'ADAPTER_ENVIRONMENT_UNSUPPORTED' });
    expect(registry.list({ kind: 'solver' }).map(({ id }) => id)).toEqual(['linear']);
    expect(diagnostics).toHaveBeenCalledWith(
      expect.objectContaining({ code: 'ADAPTER_DUPLICATE_ID' }),
    );
  });

  it('rejects malformed public fields before initialization and preserves the registry snapshot', async () => {
    const initialize = vi.fn();
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });

    await expect(
      registry.register(
        solver('invalid', {
          manifest: {
            ...solver('invalid').manifest,
            execution: 'same-realm-sandbox' as never,
          },
          initialize,
        }),
      ),
    ).rejects.toMatchObject({ code: 'ADAPTER_MANIFEST_INVALID' });
    expect(initialize).not.toHaveBeenCalled();
    expect(registry.list()).toEqual([]);
  });

  it('uses configured defaults, otherwise rejects ambiguous resolution deterministically', async () => {
    const ambiguous = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await ambiguous.register(solver('zeta'));
    await ambiguous.register(solver('alpha'));

    expect(() => ambiguous.resolve('solver')).toThrowError(
      expect.objectContaining({
        code: 'ADAPTER_AMBIGUOUS',
        diagnostic: expect.objectContaining({
          details: { candidates: ['alpha', 'zeta'], kind: 'solver' },
        }),
      }),
    );

    const configured = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      defaults: { solver: 'zeta' },
    });
    await configured.register(solver('zeta'));
    await configured.register(solver('alpha'));
    expect(configured.resolve('solver')).toMatchObject({
      manifest: { id: 'zeta' },
      reason: 'configured-default',
    });
  });

  it('snapshots configured defaults once and deeply freezes diagnostic details', async () => {
    let reads = 0;
    const defaults = {
      get solver() {
        reads += 1;
        return reads === 1 ? 'alpha' : 'zeta';
      },
    };
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      defaults,
    });
    await registry.register(solver('alpha'));
    await registry.register(solver('zeta'));

    expect(registry.resolve('solver')).toMatchObject({ manifest: { id: 'alpha' } });
    expect(registry.resolve('solver')).toMatchObject({ manifest: { id: 'alpha' } });
    expect(reads).toBe(1);

    const ambiguous = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await ambiguous.register(solver('alpha'));
    await ambiguous.register(solver('zeta'));
    let error: AdapterSdkError | undefined;
    try {
      ambiguous.resolve('solver');
    } catch (cause) {
      error = cause as AdapterSdkError;
    }
    expect(error).toBeDefined();
    expect(Object.isFrozen(error!.diagnostic.details)).toBe(true);
    expect(
      Object.isFrozen((error!.diagnostic.details as { candidates: readonly string[] }).candidates),
    ).toBe(true);
  });

  it('reads manifest fields once and rejects accessor-backed manifest arrays without invoking them', async () => {
    const reads = { id: 0, apiVersion: 0, kind: 0 };
    const capabilityGetter = vi.fn(() => 'solve');
    const capabilities: string[] = [];
    Object.defineProperty(capabilities, '0', {
      enumerable: true,
      configurable: true,
      get: capabilityGetter,
    });
    capabilities.length = 1;
    const manifest = {
      get id() {
        reads.id += 1;
        return 'hostile';
      },
      get apiVersion() {
        reads.apiVersion += 1;
        return '1.0' as const;
      },
      get kind() {
        reads.kind += 1;
        return 'solver' as const;
      },
      environments: ['browser'] as const,
      execution: 'trusted-main' as const,
      priority: 0,
      capabilities,
    };
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });

    await expect(
      registry.register({
        manifest,
        implementation: { invoke: async () => null },
      }),
    ).rejects.toMatchObject({ code: 'ADAPTER_MANIFEST_INVALID' });
    expect(reads).toEqual({ id: 1, apiVersion: 1, kind: 1 });
    expect(capabilityGetter).not.toHaveBeenCalled();
  });

  it('releases registrations in reverse order and makes unregister/dispose idempotent', async () => {
    const released: string[] = [];
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    const unregisterAlpha = await registry.register(
      solver('alpha', { dispose: () => void released.push('alpha') }),
    );
    await registry.register(solver('zeta', { dispose: () => void released.push('zeta') }));
    await registry.register(solver('gamma', { dispose: () => void released.push('gamma') }));

    await unregisterAlpha();
    await unregisterAlpha();
    expect(released).toEqual(['alpha']);

    expect(await registry.dispose()).toEqual([]);
    expect(await registry.dispose()).toEqual([]);
    expect(released).toEqual(['alpha', 'gamma', 'zeta']);
    expect(() => registry.resolve('solver', { id: 'zeta' })).toThrowError(AdapterSdkError);
  });

  it('compensates failed initialization and aggregates public disposal diagnostics', async () => {
    const cleanup = vi.fn(async () => {
      throw new Error('cleanup failed');
    });
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });

    await expect(
      registry.register(
        solver('broken', {
          initialize: async () => {
            throw new Error('initialize failed');
          },
          dispose: cleanup,
        }),
      ),
    ).rejects.toMatchObject({
      code: 'ADAPTER_INITIALIZATION_FAILED',
      diagnostics: [
        expect.objectContaining({ code: 'ADAPTER_INITIALIZATION_FAILED' }),
        expect.objectContaining({ code: 'ADAPTER_DISPOSE_FAILED' }),
      ],
    });
    expect(cleanup).toHaveBeenCalledTimes(1);
    expect(registry.list()).toEqual([]);
  });
});
