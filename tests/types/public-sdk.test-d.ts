import {
  createTemplateModuleRegistry,
  executeTemplateModulePipeline,
  type TemplateModule,
  type TemplateModulePipelineOutput,
} from '../../src/sdk';

const registry = createTemplateModuleRegistry({
  apiVersion: '1.0',
  environment: 'worker',
});

const templateModule: TemplateModule = {
  manifest: {
    id: 'example.module',
    apiVersion: '1.0',
    environments: ['worker'],
    priority: 0,
  },
  recognize: (node) => node.type === 'example.module',
  transform: (node) => node.data,
  resolve: (ir) => ir,
  layout: (ir) => ir,
  paint: () => [],
};

const unregister: () => Promise<readonly unknown[]> = registry.register(templateModule);
const output: Promise<TemplateModulePipelineOutput> = executeTemplateModulePipeline(registry, {
  document: {
    requiredModules: ['example.module'],
    nodes: [{ id: 'example.node', type: 'example.module', data: null }],
  },
  signal: new AbortController().signal,
});

void unregister;
void output;

// @ts-expect-error modules cannot add arbitrary rendering stages
templateModule.measure = () => null;
