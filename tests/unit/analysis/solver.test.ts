import { describe, expect, it, vi } from 'vitest';
import { compileSolverModel, runSolver } from '../../../src/analysis/solver';
import { createAdapterRegistry } from '../../../src/sdk/adapters';

const request = {
  documentId: 'document-1',
  revision: 'revision-1',
  objective: {
    goal: 'minimize' as const,
    address: { sheetId: 'sheet-1', row: 0, column: 2 },
  },
  variables: [
    {
      id: 'x',
      address: { sheetId: 'sheet-1', row: 0, column: 0 },
      minimum: 0,
      maximum: 10,
    },
  ],
  constraints: [
    {
      id: 'capacity',
      address: { sheetId: 'sheet-1', row: 0, column: 1 },
      operator: '<=' as const,
      value: 100,
    },
  ],
};

describe('solver adapter boundary', () => {
  it('compiles a bounded immutable problem IR without a document or controller', () => {
    const model = compileSolverModel(request);

    expect(model).toEqual(request);
    expect(Object.isFrozen(model)).toBe(true);
    expect(Object.isFrozen(model.variables)).toBe(true);
    expect(JSON.stringify(model)).not.toContain('controller');
  });

  it('rejects invalid and excessive models before adapter execution', () => {
    expect(() =>
      compileSolverModel({
        ...request,
        variables: [{ ...request.variables[0]!, minimum: 2, maximum: 1 }],
      }),
    ).toThrow(/bounds/u);
    expect(() =>
      compileSolverModel({
        ...request,
        variables: Array.from({ length: 10_001 }, (_, index) => ({
          id: `v-${index}`,
          address: { sheetId: 'sheet-1', row: index, column: 0 },
        })),
      }),
    ).toThrow(/limit/u);
    expect(() =>
      compileSolverModel({
        ...request,
        variables: [{ ...request.variables[0]!, integer: 'yes' as never }],
      }),
    ).toThrow(/integer/u);
    expect(() =>
      compileSolverModel({
        ...request,
        objective: { ...request.objective, targetValue: 10 },
      }),
    ).toThrow(/targetValue/u);
  });

  it('runs only through an isolated-worker adapter and validates its result', async () => {
    const invoke = vi.fn(async () => ({
      status: 'optimal',
      objectiveValue: 12,
      candidates: [{ variableId: 'x', value: 4 }],
      residuals: [{ constraintId: 'capacity', value: 0 }],
    }));
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      isolatedWorkerTransport: {
        invoke,
        terminate: vi.fn(async () => undefined),
      },
    });
    await registry.register({
      manifest: {
        id: 'worker-solver',
        apiVersion: '1.0',
        kind: 'solver',
        environments: ['browser'],
        execution: 'isolated-worker',
        priority: 0,
        capabilities: ['solve'],
      },
      descriptor: { workerId: 'solver-worker' },
    });

    await expect(
      runSolver(registry, compileSolverModel(request), {
        signal: new AbortController().signal,
      }),
    ).resolves.toEqual({
      status: 'optimal',
      objectiveValue: 12,
      candidates: [{ variableId: 'x', value: 4 }],
      residuals: [{ constraintId: 'capacity', value: 0 }],
    });
    expect(invoke).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'solver',
        capability: 'solve',
        input: expect.not.objectContaining({ controller: expect.anything() }),
      }),
      expect.any(AbortSignal),
    );
  });

  it('rejects trusted-main solver execution', async () => {
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
    });
    await registry.register({
      manifest: {
        id: 'main-solver',
        apiVersion: '1.0',
        kind: 'solver',
        environments: ['browser'],
        execution: 'trusted-main',
        priority: 0,
        capabilities: ['solve'],
      },
      implementation: { invoke: vi.fn() },
    });

    await expect(
      runSolver(registry, compileSolverModel(request), {
        signal: new AbortController().signal,
      }),
    ).rejects.toThrow(/isolated-worker/u);
  });

  it('rejects incomplete, out-of-bounds, non-integral, and status-inconsistent results', async () => {
    const cases = [
      {
        status: 'optimal',
        objectiveValue: 12,
        candidates: [],
        residuals: [{ constraintId: 'capacity', value: 0 }],
      },
      {
        status: 'optimal',
        objectiveValue: 12,
        candidates: [{ variableId: 'x', value: 11 }],
        residuals: [{ constraintId: 'capacity', value: 0 }],
      },
      {
        status: 'optimal',
        candidates: [{ variableId: 'x', value: 4 }],
        residuals: [{ constraintId: 'capacity', value: 0 }],
      },
      {
        status: 'cancelled',
        objectiveValue: 12,
        candidates: [{ variableId: 'x', value: 4 }],
        residuals: [{ constraintId: 'capacity', value: 0 }],
      },
    ];

    for (const output of cases) {
      const registry = createAdapterRegistry({
        apiVersion: '1.0',
        environment: 'browser',
        isolatedWorkerTransport: {
          invoke: vi.fn(async () => output),
          terminate: vi.fn(async () => undefined),
        },
      });
      await registry.register({
        manifest: {
          id: 'strict-worker-solver',
          apiVersion: '1.0',
          kind: 'solver',
          environments: ['browser'],
          execution: 'isolated-worker',
          priority: 0,
          capabilities: ['solve'],
        },
        descriptor: { workerId: 'strict-solver-worker' },
      });
      await expect(
        runSolver(registry, compileSolverModel(request), {
          signal: new AbortController().signal,
        }),
      ).rejects.toThrow();
    }

    const integerModel = compileSolverModel({
      ...request,
      variables: [{ ...request.variables[0]!, integer: true }],
    });
    const registry = createAdapterRegistry({
      apiVersion: '1.0',
      environment: 'browser',
      isolatedWorkerTransport: {
        invoke: vi.fn(async () => ({
          status: 'feasible',
          objectiveValue: 12,
          candidates: [{ variableId: 'x', value: 4.5 }],
          residuals: [{ constraintId: 'capacity', value: 0 }],
        })),
        terminate: vi.fn(async () => undefined),
      },
    });
    await registry.register({
      manifest: {
        id: 'integer-worker-solver',
        apiVersion: '1.0',
        kind: 'solver',
        environments: ['browser'],
        execution: 'isolated-worker',
        priority: 0,
        capabilities: ['solve'],
      },
      descriptor: { workerId: 'integer-solver-worker' },
    });
    await expect(
      runSolver(registry, integerModel, {
        signal: new AbortController().signal,
      }),
    ).rejects.toThrow();
  });
});
