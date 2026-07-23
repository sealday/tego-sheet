import { describe, expect, it } from 'vitest';
import { createAdapterRegistryKernel } from '../../../src/extensions/kernel/registry';
import type { KernelRegistration } from '../../../src/extensions/kernel/capabilities';

describe('formula function kernel capability', () => {
  it('registers and resolves a typed formula function through the F5 kernel', async () => {
    const kernel = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'worker' });
    const registration: KernelRegistration<'formula-function'> = {
      manifest: {
        id: 'double',
        kind: 'formula-function',
        apiVersion: '1.0',
        environments: ['worker'],
      },
      implementation: {
        name: 'DOUBLE',
        parameters: { minimum: 1, maximum: 1 },
        returns: 'number',
        volatility: 'stable',
        mode: 'sync',
        evaluate: ([value]) => ({
          type: 'number',
          value: value?.type === 'number' ? value.value * 2 : 0,
        }),
      },
    };
    await kernel.register(registration);
    expect(kernel.resolve('formula-function', { id: 'double', environment: 'worker' }).name).toBe(
      'DOUBLE',
    );
  });
});
