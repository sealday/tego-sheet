import { describe, expect, it, vi } from 'vitest';
import {
  TemplateModuleSdkError,
  createTemplateModuleRegistry,
  executeTemplateModulePipeline,
  type TemplateModule,
  type TemplateModuleNode,
} from '../../../src/sdk';

const node = (id: string, type: string, data = {}): TemplateModuleNode => ({
  id,
  type,
  data,
});

function module(
  id: string,
  options: Partial<TemplateModule> & { readonly priority?: number } = {},
): TemplateModule {
  const withStage = (ir: unknown, stage: string): { readonly stage: string } => ({
    ...(ir as Readonly<Record<string, unknown>>),
    stage,
  });
  return {
    manifest: {
      id,
      apiVersion: '1.0',
      environments: ['node'],
      priority: options.priority ?? 0,
    },
    recognize: options.recognize ?? ((candidate) => candidate.type === id),
    transform: options.transform ?? ((candidate) => ({ stage: 'transform', id: candidate.id })),
    resolve: options.resolve ?? ((ir) => withStage(ir, 'resolve')),
    layout: options.layout ?? ((ir) => withStage(ir, 'layout')),
    paint:
      options.paint ??
      (() => [
        {
          kind: 'fill-rect',
          rect: { x: 0, y: 0, width: 10, height: 10 },
          color: '#000000',
        },
      ]),
    ...(options.dispose === undefined ? {} : { dispose: options.dispose }),
  };
}

function codes(error: unknown): readonly string[] {
  expect(error).toBeInstanceOf(TemplateModuleSdkError);
  return (error as TemplateModuleSdkError).diagnostics.map((diagnostic) => diagnostic.code);
}

describe('template module SDK', () => {
  it('negotiates API and environment compatibility and lists modules stably', () => {
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.4',
      environment: 'node',
    });
    registry.register(module('zeta', { priority: 1 }));
    registry.register(module('alpha', { priority: 2 }));
    registry.register(module('beta', { priority: 2 }));

    expect(registry.list().map(({ manifest }) => manifest.id)).toEqual(['alpha', 'beta', 'zeta']);
    expect(() =>
      registry.register({
        ...module('future'),
        manifest: {
          id: 'future',
          apiVersion: '2.0',
          environments: ['node'],
          priority: 0,
        },
      }),
    ).toThrow(TemplateModuleSdkError);
    expect(() =>
      registry.register({
        ...module('browser-only'),
        manifest: {
          id: 'browser-only',
          apiVersion: '1.0',
          environments: ['browser'],
          priority: 0,
        },
      }),
    ).toThrow(TemplateModuleSdkError);
  });

  it('runs recognize through paint in finite deterministic stage order', async () => {
    const calls: string[] = [];
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(
      module('widget', {
        recognize(candidate) {
          calls.push(`recognize:${candidate.id}`);
          return candidate.type === 'widget';
        },
        transform(candidate) {
          calls.push(`transform:${candidate.id}`);
          return { value: 1 };
        },
        resolve(ir) {
          calls.push(`resolve:${String((ir as { value: number }).value)}`);
          return { value: 2 };
        },
        layout(ir) {
          calls.push(`layout:${String((ir as { value: number }).value)}`);
          return { value: 3 };
        },
        paint(ir) {
          calls.push(`paint:${String((ir as { value: number }).value)}`);
          return [
            {
              kind: 'text',
              text: 'ok',
              x: 1,
              y: 2,
              maxWidth: 30,
              fontFamily: 'sans-serif',
              fontSize: 12,
              color: '#000',
              horizontalAlign: 'left',
            },
          ];
        },
      }),
    );

    const result = await executeTemplateModulePipeline(registry, {
      document: { requiredModules: ['widget'], nodes: [node('n1', 'widget')] },
      signal: new AbortController().signal,
    });

    expect(calls).toEqual(['recognize:n1', 'transform:n1', 'resolve:1', 'layout:2', 'paint:3']);
    expect(result.nodes[0]).toMatchObject({ nodeId: 'n1', moduleId: 'widget' });
    expect(result.commands).toHaveLength(1);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.commands[0])).toBe(true);
  });

  it('rejects conflicting node ownership before transform or output', async () => {
    const transform = vi.fn(() => ({}));
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(module('one', { recognize: () => true, transform }));
    registry.register(module('two', { recognize: () => true, transform }));

    await expect(
      executeTemplateModulePipeline(registry, {
        document: { requiredModules: [], nodes: [node('same', 'shared')] },
        signal: new AbortController().signal,
      }),
    ).rejects.toSatisfy((error: unknown) =>
      codes(error).includes('TEMPLATE_NODE_OWNERSHIP_CONFLICT'),
    );
    expect(transform).not.toHaveBeenCalled();
  });

  it('blocks all output when a required module is missing', async () => {
    const paint = vi.fn(() => []);
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(module('present', { paint }));

    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['present', 'missing'],
          nodes: [node('n1', 'present')],
        },
        signal: new AbortController().signal,
      }),
    ).rejects.toSatisfy((error: unknown) => codes(error).includes('TEMPLATE_MODULE_MISSING'));
    expect(paint).not.toHaveBeenCalled();
  });

  it('snapshots and freezes node and IR values across every module boundary', async () => {
    const sourceData = { nested: { value: 1 } };
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(
      module('widget', {
        transform(candidate) {
          expect(Object.isFrozen(candidate)).toBe(true);
          expect(Object.isFrozen(candidate.data)).toBe(true);
          return candidate.data;
        },
        resolve(ir) {
          expect(Object.isFrozen(ir)).toBe(true);
          return ir;
        },
        layout(ir) {
          expect(Object.isFrozen(ir)).toBe(true);
          return ir;
        },
      }),
    );

    const pending = executeTemplateModulePipeline(registry, {
      document: {
        requiredModules: ['widget'],
        nodes: [node('n1', 'widget', sourceData)],
      },
      signal: new AbortController().signal,
    });
    sourceData.nested.value = 99;
    const result = await pending;

    expect(result.nodes[0]?.ir).toEqual({ nested: { value: 1 } });
    const firstNode = result.nodes[0];
    if (firstNode === undefined) throw new Error('expected transformed node');
    expect(Object.isFrozen((firstNode.ir as { nested: object }).nested)).toBe(true);
  });

  it('enforces node, output, and command budgets', async () => {
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(
      module('widget', {
        transform: () => ({ text: 'larger-than-budget' }),
        paint: () => [
          {
            kind: 'fill-rect',
            rect: { x: 0, y: 0, width: 1, height: 1 },
            color: '#000',
          },
          {
            kind: 'fill-rect',
            rect: { x: 1, y: 1, width: 1, height: 1 },
            color: '#fff',
          },
        ],
      }),
    );

    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget'), node('n2', 'widget')],
        },
        signal: new AbortController().signal,
        limits: { maximumNodes: 1 },
      }),
    ).rejects.toSatisfy((error: unknown) => codes(error).includes('TEMPLATE_NODE_LIMIT_EXCEEDED'));
    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget')],
        },
        signal: new AbortController().signal,
        limits: { maximumOutputBytes: 5 },
      }),
    ).rejects.toSatisfy((error: unknown) =>
      codes(error).includes('TEMPLATE_OUTPUT_LIMIT_EXCEEDED'),
    );
    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget')],
        },
        signal: new AbortController().signal,
        limits: { maximumCommands: 1 },
      }),
    ).rejects.toSatisfy((error: unknown) =>
      codes(error).includes('TEMPLATE_COMMAND_LIMIT_EXCEEDED'),
    );
  });

  it('checks elapsed-time budget and AbortSignal between every stage', async () => {
    let now = 0;
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(
      module('widget', {
        recognize: () => {
          now = 20;
          return true;
        },
      }),
    );

    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget')],
        },
        signal: new AbortController().signal,
        limits: { maximumMilliseconds: 10 },
        clock: () => now,
      }),
    ).rejects.toSatisfy((error: unknown) => codes(error).includes('TEMPLATE_TIME_LIMIT_EXCEEDED'));

    const controller = new AbortController();
    controller.abort('stop');
    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget')],
        },
        signal: controller.signal,
      }),
    ).rejects.toSatisfy((error: unknown) => codes(error).includes('TEMPLATE_PIPELINE_ABORTED'));
  });

  it('rejects malformed or unsupported draw commands', async () => {
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    registry.register(
      module('widget', {
        paint: () =>
          [
            { kind: 'script', source: 'alert(1)' },
            { kind: 'fill-rect', color: '#000' },
          ] as never,
      }),
    );

    await expect(
      executeTemplateModulePipeline(registry, {
        document: {
          requiredModules: ['widget'],
          nodes: [node('n1', 'widget')],
        },
        signal: new AbortController().signal,
      }),
    ).rejects.toSatisfy((error: unknown) => codes(error).includes('TEMPLATE_DRAW_COMMAND_INVALID'));
  });

  it('disposes modules once, unregisters independently, and rejects work after disposal', async () => {
    const firstDispose = vi.fn();
    const secondDispose = vi.fn();
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    const unregister = registry.register(module('first', { dispose: firstDispose }));
    registry.register(module('second', { dispose: secondDispose }));

    await unregister();
    await unregister();
    expect(firstDispose).toHaveBeenCalledTimes(1);
    await registry.dispose();
    await registry.dispose();
    expect(firstDispose).toHaveBeenCalledTimes(1);
    expect(secondDispose).toHaveBeenCalledTimes(1);
    expect(() => registry.register(module('later'))).toThrow(TemplateModuleSdkError);
  });

  it('does not dispose a module twice when its unregister handle runs after registry disposal', async () => {
    const dispose = vi.fn();
    const registry = createTemplateModuleRegistry({
      apiVersion: '1.0',
      environment: 'node',
    });
    const unregister = registry.register(module('owned', { dispose }));

    await registry.dispose();
    await unregister();

    expect(dispose).toHaveBeenCalledTimes(1);
  });

  it('waits for an active pipeline to settle before disposing its modules', async () => {
    let releaseRecognition!: (value: boolean) => void;
    const recognition = new Promise<boolean>((resolve) => {
      releaseRecognition = resolve;
    });
    const dispose = vi.fn();
    const registry = createTemplateModuleRegistry({ apiVersion: '1.0', environment: 'node' });
    registry.register(module('owned', { recognize: () => recognition, dispose }));
    const pipeline = executeTemplateModulePipeline(registry, {
      document: { requiredModules: ['owned'], nodes: [node('n1', 'owned')] },
      signal: new AbortController().signal,
    });

    const disposing = registry.dispose();
    await Promise.resolve();
    expect(dispose).not.toHaveBeenCalled();
    releaseRecognition(true);
    await pipeline;
    await disposing;

    expect(dispose).toHaveBeenCalledTimes(1);
  });
});
