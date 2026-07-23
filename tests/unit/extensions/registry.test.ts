import { describe, expect, it, vi } from 'vitest';
import type {
  BuiltInCellTypeDefinition,
  KernelRegistration,
} from '../../../src/extensions/kernel/capabilities';
import { createAdapterRegistryKernel } from '../../../src/extensions/kernel/registry';

function definition(id: string): BuiltInCellTypeDefinition<string> {
  return {
    id,
    schemaVersion: 1,
    validate: (value): value is string => typeof value === 'string',
    describe: (value) => ({
      formattedText: value,
      accessibilityLabel: value,
      role: 'text',
    }),
    toFormulaScalar: (value) => value,
  };
}

function registration(
  id: string,
  options: Partial<KernelRegistration<'cell-type'>['manifest']> = {},
): KernelRegistration<'cell-type'> {
  return {
    manifest: {
      id,
      apiVersion: '1.0',
      kind: 'cell-type',
      environments: ['browser'],
      ...options,
    },
    implementation: definition(id),
  };
}

describe('AdapterRegistryKernel resolution', () => {
  it('resolves an exact kind and id in a supported environment in O(1) lookup semantics', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.2', environment: 'browser' });
    const checkbox = registration('checkbox');
    await registry.register(checkbox);

    expect(registry.resolve('cell-type', { id: 'checkbox', environment: 'browser' })).toBe(
      checkbox.implementation,
    );
  });

  it('rejects a duplicate kind and id without initializing the duplicate', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    const initialize = vi.fn();
    await registry.register(registration('checkbox'));

    await expect(
      registry.register({ ...registration('checkbox'), initialize }),
    ).rejects.toMatchObject({
      code: 'EXTENSION_DUPLICATE_ID',
    });
    expect(initialize).not.toHaveBeenCalled();
  });

  it('does not fall back to a different id or unsupported environment', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    await registry.register(registration('checkbox'));

    expect(() =>
      registry.resolve('cell-type', { id: 'check', environment: 'browser' }),
    ).toThrowError(expect.objectContaining({ code: 'EXTENSION_NOT_FOUND' }));
    expect(() =>
      registry.resolve('cell-type', { id: 'checkbox', environment: 'worker' }),
    ).toThrowError(expect.objectContaining({ code: 'EXTENSION_ENVIRONMENT_UNSUPPORTED' }));
  });

  it('requires an unambiguous default and is independent of registration order', async () => {
    for (const ids of [
      ['zeta', 'alpha'],
      ['alpha', 'zeta'],
    ]) {
      const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
      for (const id of ids) await registry.register(registration(id));

      expect(() => registry.resolve('cell-type', { environment: 'browser' })).toThrowError(
        expect.objectContaining({
          code: 'EXTENSION_AMBIGUOUS',
          diagnostic: expect.objectContaining({
            details: { candidates: ['alpha', 'zeta'], kind: 'cell-type' },
          }),
        }),
      );
    }
  });

  it('uses the only environment-compatible default and reports zero candidates', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    const browser = registration('browser', { environments: ['browser'] });
    await registry.register(browser);
    await registry.register(registration('worker', { environments: ['worker'] }));

    expect(registry.resolve('cell-type', { environment: 'browser' })).toBe(browser.implementation);
    expect(() => registry.resolve('cell-type', { environment: 'node' })).toThrowError(
      expect.objectContaining({ code: 'EXTENSION_ENVIRONMENT_UNSUPPORTED' }),
    );

    const empty = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'node' });
    expect(() => empty.resolve('cell-type', { environment: 'node' })).toThrowError(
      expect.objectContaining({ code: 'EXTENSION_NOT_FOUND' }),
    );
  });

  it('lists frozen manifest snapshots in stable kind and id order', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    await registry.register(registration('zeta'));
    await registry.register(registration('alpha'));

    const listed = registry.list();
    expect(listed.map(({ kind, id }) => `${kind}/${id}`)).toEqual([
      'cell-type/alpha',
      'cell-type/zeta',
    ]);
    expect(Object.isFrozen(listed)).toBe(true);
    expect(Object.isFrozen(listed[0])).toBe(true);
  });
});

describe('AdapterRegistryKernel compatibility and lifecycle', () => {
  it('rejects malformed runtime manifests with a stable diagnostic instead of leaking a TypeError', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });

    await expect(
      registry.register({
        ...registration('checkbox'),
        manifest: {
          ...registration('checkbox').manifest,
          environments: null as never,
        },
      }),
    ).rejects.toMatchObject({ code: 'EXTENSION_MANIFEST_INVALID' });
  });

  it.each([
    ['2.0', '1.9'],
    ['1.3', '1.2'],
  ])('rejects unsupported registration API %s for package API %s', async (requested, supported) => {
    const registry = createAdapterRegistryKernel({
      apiVersion: supported as `${number}.${number}`,
      environment: 'browser',
    });

    await expect(
      registry.register(
        registration('checkbox', { apiVersion: requested as `${number}.${number}` }),
      ),
    ).rejects.toMatchObject({ code: 'EXTENSION_API_INCOMPATIBLE' });
  });

  it('accepts the supported major and an equal or older minor version', async () => {
    const registry = createAdapterRegistryKernel({ apiVersion: '1.2', environment: 'browser' });
    await registry.register(registration('old', { apiVersion: '1.0' }));
    await registry.register(registration('current', { apiVersion: '1.2' }));

    expect(registry.list()).toHaveLength(2);
  });

  it('compensates failed initialization and never publishes the partial registration', async () => {
    const events: string[] = [];
    const registry = createAdapterRegistryKernel({
      apiVersion: '1.0',
      environment: 'browser',
    });

    await expect(
      registry.register({
        ...registration('broken'),
        initialize: () => {
          events.push('initialize');
          throw new Error('cannot start');
        },
        dispose: () => {
          events.push('dispose');
        },
      }),
    ).rejects.toMatchObject({
      code: 'EXTENSION_INITIALIZE_FAILED',
      diagnostics: [
        expect.objectContaining({
          code: 'EXTENSION_INITIALIZE_FAILED',
          domain: 'extension',
          stage: 'execute',
        }),
      ],
    });
    expect(events).toEqual(['initialize', 'dispose']);
    expect(registry.list()).toEqual([]);
  });

  it('unregisters once and makes repeated handles and global disposal idempotent', async () => {
    const dispose = vi.fn();
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    const unregister = await registry.register({ ...registration('checkbox'), dispose });

    await expect(unregister()).resolves.toEqual([]);
    await expect(unregister()).resolves.toEqual([]);
    await expect(registry.dispose()).resolves.toEqual([]);
    await expect(registry.dispose()).resolves.toEqual([]);
    expect(dispose).toHaveBeenCalledTimes(1);
    expect(registry.list()).toEqual([]);
  });

  it('continues cleanup and aggregates multiple disposal failures in stable order', async () => {
    const events: string[] = [];
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    for (const id of ['zeta', 'alpha']) {
      await registry.register({
        ...registration(id),
        dispose: () => {
          events.push(id);
          throw new Error(`${id} failed`);
        },
      });
    }

    const diagnostics = await registry.dispose();

    expect(events).toEqual(['alpha', 'zeta']);
    expect(diagnostics.map(({ code, location }) => [code, location?.adapterId])).toEqual([
      ['EXTENSION_DISPOSE_FAILED', 'alpha'],
      ['EXTENSION_DISPOSE_FAILED', 'zeta'],
    ]);
    await expect(registry.dispose()).resolves.toEqual([]);
  });

  it('aborts the context before disposing a registration', async () => {
    let signal: AbortSignal | undefined;
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    const unregister = await registry.register({
      ...registration('checkbox'),
      initialize: (context) => {
        signal = context.signal;
      },
      dispose: () => {
        expect(signal?.aborted).toBe(true);
      },
    });

    expect(signal?.aborted).toBe(false);
    await unregister();
  });

  it('aborts and waits for in-flight initialization before global disposal resolves', async () => {
    let finishInitialization!: () => void;
    let signal: AbortSignal | undefined;
    const initializationGate = new Promise<void>((resolve) => {
      finishInitialization = resolve;
    });
    const dispose = vi.fn();
    const registry = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'browser' });
    const registrationPromise = registry.register({
      ...registration('slow'),
      initialize: async (context) => {
        signal = context.signal;
        await initializationGate;
      },
      dispose,
    });

    await vi.waitFor(() => expect(signal).toBeDefined());
    let disposalSettled = false;
    const disposalPromise = registry.dispose().then((diagnostics) => {
      disposalSettled = true;
      return diagnostics;
    });
    await Promise.resolve();

    expect(signal?.aborted).toBe(true);
    expect(disposalSettled).toBe(false);

    finishInitialization();
    await expect(registrationPromise).rejects.toMatchObject({
      code: 'EXTENSION_REGISTRY_DISPOSED',
    });
    await expect(disposalPromise).resolves.toEqual([]);
    expect(dispose).toHaveBeenCalledTimes(1);
  });
});
