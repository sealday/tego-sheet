import { describe, expect, it } from 'vitest';
import { createAdapterRegistryKernel } from '../../../src/extensions/kernel/registry';
import type { KernelRegistration } from '../../../src/extensions/kernel/capabilities';
import { createFormulaEngine } from '../../../src/formula/evaluator';
import {
  createFormulaFunctionRegistry,
  registerKernelFormulaFunctions,
} from '../../../src/formula/function-registry';
import { formulaDocument } from './helpers';

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

    const functions = createFormulaFunctionRegistry();
    const unregister = registerKernelFormulaFunctions(functions, kernel, 'worker');
    const engine = createFormulaEngine({ functions });
    const document = formulaDocument([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: [{ row: 0, column: 0, input: { type: 'formula', source: '=DOUBLE(4)' } }],
      },
    ]);
    expect(
      engine
        .recalculate(engine.compile(document), [], {
          locale: 'en-US',
          timeZone: 'UTC',
          dateSystem: 'excel-1900',
          clock: { now: () => 0 },
          tick: 0,
          functionRegistryVersion: functions.version,
        })
        .values.get('sheet-1!A1'),
    ).toEqual({ type: 'number', value: 8 });
    unregister();
    expect(functions.resolve('DOUBLE')).toBeUndefined();
  });

  it('rolls back copied definitions when one kernel function cannot be registered', async () => {
    const kernel = createAdapterRegistryKernel({ apiVersion: '1.0', environment: 'worker' });
    const implementation: KernelRegistration<'formula-function'>['implementation'] = {
      name: 'DUPLICATE',
      parameters: { minimum: 0, maximum: 0 },
      returns: 'number',
      volatility: 'stable',
      mode: 'sync',
      evaluate: () => ({ type: 'number', value: 1 }),
    };
    for (const id of ['first', 'second']) {
      await kernel.register({
        manifest: {
          id,
          kind: 'formula-function',
          apiVersion: '1.0',
          environments: ['worker'],
        },
        implementation,
      });
    }
    const functions = createFormulaFunctionRegistry();

    expect(() => registerKernelFormulaFunctions(functions, kernel, 'worker')).toThrow(
      /already registered/u,
    );
    expect(functions.resolve('DUPLICATE')).toBeUndefined();
  });
});
