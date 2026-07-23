import type {
  AdvancedCompileOptions,
  RenderEnvironment,
  RenderRequest,
  ResourceResolver,
  ResolvedResourceStore,
  TemplateBinding,
} from '../../src';
import type { KernelCapabilities } from '../../src/extensions/kernel/capabilities';

declare const resolver: ResourceResolver;
declare const store: ResolvedResourceStore;
declare const environment: RenderEnvironment;
declare const request: RenderRequest;
declare const advanced: AdvancedCompileOptions;

const capability: KernelCapabilities['resource-resolver'] = resolver;
const binding: TemplateBinding = {
  id: 'repeat-columns' as never,
  type: 'repeat-columns',
  range: {
    sheetId: 'sheet-1' as never,
    start: { row: 0, column: 0 },
    end: { row: 0, column: 0 },
  },
  source: 'labels',
  empty: 'remove',
};

void capability;
void binding;
void store.byHash;
void environment.resourceRegistry;
void request.resourceRefs;
void advanced.subtemplates;
