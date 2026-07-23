import type { JsonValue } from '../../src/document';
import type {
  BuiltInCellTypeDefinition,
  KernelRegistration,
} from '../../src/extensions/kernel/capabilities';
import type { AdapterRegistryKernel } from '../../src/extensions/kernel/registry';

declare module '../../src/extensions/kernel/capabilities' {
  interface KernelCapabilities {
    'example-number': {
      readonly parse: (value: string) => number;
    };
  }
}

declare const registry: AdapterRegistryKernel;
declare const cellType: BuiltInCellTypeDefinition<JsonValue>;

const cellResolution: BuiltInCellTypeDefinition<JsonValue> = registry.resolve('cell-type', {
  id: 'checkbox',
  environment: 'browser',
});
const numberResolution: { readonly parse: (value: string) => number } = registry.resolve(
  'example-number',
  { environment: 'node' },
);

const numberRegistration: KernelRegistration<'example-number'> = {
  manifest: {
    id: 'number',
    apiVersion: '1.0',
    kind: 'example-number',
    environments: ['node'],
  },
  implementation: { parse: Number },
};

void cellResolution;
void numberResolution;
void numberRegistration;

const invalidRegistration: KernelRegistration<'example-number'> = {
  manifest: {
    id: 'invalid',
    apiVersion: '1.0',
    kind: 'example-number',
    environments: ['node'],
  },
  // @ts-expect-error capability implementations keep the exact declaration-merged type
  implementation: cellType,
};

void invalidRegistration;

// @ts-expect-error unregistered capability kinds cannot be resolved
registry.resolve('missing-kind', { environment: 'browser' });
