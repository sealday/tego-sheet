import {
  createResourceResolverRegistry,
  type ResourceResolver,
  type ResourceResolverRegistry,
} from './resource-pipeline';

/** Runtime environment accepted by F5 resource capabilities. */
export type ResourceKernelEnvironment = 'browser' | 'node' | 'worker';

/** Narrow structural view of the F5 kernel required by the resource bridge. */
export interface ResourceCapabilityRegistry {
  /** Resolves one explicit resource capability. */
  resolve(
    kind: 'resource-resolver',
    query: { readonly id: string; readonly environment: ResourceKernelEnvironment },
  ): ResourceResolver;
}

/** Builds a deterministic resource registry from explicit F5 capability IDs. */
export function createResourceResolverRegistryFromKernel(
  kernel: ResourceCapabilityRegistry,
  resolverIds: readonly string[],
  environment: ResourceKernelEnvironment,
): ResourceResolverRegistry {
  const unique = new Set<string>();
  const resolvers = resolverIds.map((id) => {
    if (unique.has(id)) throw new TypeError(`Duplicate resource resolver capability ${id}`);
    unique.add(id);
    return kernel.resolve('resource-resolver', { id, environment });
  });
  return createResourceResolverRegistry(resolvers);
}
