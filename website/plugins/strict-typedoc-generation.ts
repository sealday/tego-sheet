import { rm } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import type { LoadContext, Plugin } from '@docusaurus/types';
import type { PluginOptions as DocusaurusTypeDocOptions } from 'docusaurus-plugin-typedoc';
import {
  Application,
  PackageJsonReader,
  TSConfigReader,
  TypeDocReader,
  type TypeDocOptions,
} from 'typedoc';

const requiredPlugins = ['typedoc-plugin-markdown', 'typedoc-docusaurus-theme'];

interface TypeDocGenerationContext {
  siteDir: string;
  siteConfig?: {
    presets?: readonly unknown[];
  };
}

type TypeDocGenerationOptions = TypeDocOptions & DocusaurusTypeDocOptions;

interface StrictTypeDocLogger {
  errorCount: number;
  validationWarningCount: number;
  warningCount: number;
  hasErrors: () => boolean;
  hasWarnings: () => boolean;
}

export interface StrictTypeDocApplication {
  logger: StrictTypeDocLogger;
  options: {
    getValue: (name: 'treatValidationWarningsAsErrors' | 'treatWarningsAsErrors') => boolean;
  };
  convert: () => Promise<unknown | undefined>;
  generateOutputs: (project: unknown) => Promise<void>;
  validate: (project: unknown) => void;
}

interface StrictTypeDocDependencies {
  bootstrap: (options: TypeDocGenerationOptions) => Promise<StrictTypeDocApplication>;
  clearOutput: (out: string) => Promise<void>;
  projectRoot: string;
}

const defaultDependencies: StrictTypeDocDependencies = {
  async bootstrap(options) {
    const app = await Application.bootstrapWithPlugins(options, [
      new TypeDocReader(),
      new PackageJsonReader(),
      new TSConfigReader(),
    ]);

    return app as unknown as StrictTypeDocApplication;
  },
  async clearOutput(out) {
    await rm(out, { force: true, recursive: true });
  },
  projectRoot: process.cwd(),
};

const describeFailure = (phase: string, app: StrictTypeDocApplication): string =>
  `TypeDoc ${phase} failed with ${app.logger.errorCount} errors and ${app.logger.warningCount} warnings`;

const readDocsPreset = (preset: unknown): Record<string, unknown> | undefined => {
  if (!Array.isArray(preset)) return undefined;
  const options: unknown = preset[1];
  if (typeof options !== 'object' || options === null || Array.isArray(options)) return undefined;
  const docs = (options as Record<string, unknown>).docs;
  if (typeof docs !== 'object' || docs === null || Array.isArray(docs)) return undefined;
  return docs as Record<string, unknown>;
};

export function toTypeDocGenerationOptions(
  context: TypeDocGenerationContext,
  options: Record<string, unknown>,
): TypeDocGenerationOptions {
  const docs = context.siteConfig?.presets?.map(readDocsPreset).find(Boolean);
  const docsPath = typeof docs?.path === 'string' ? docs.path : './docs';
  const configuredPlugins = Array.isArray(options.plugin) ? options.plugin : [];
  const generationOptions = {
    out: './docs/api',
    docsPath: join(context.siteDir, docsPath),
    numberPrefixParser:
      typeof docs?.numberPrefixParser === 'boolean' ? docs.numberPrefixParser : true,
    ...options,
    plugin: [...new Set([...requiredPlugins, ...configuredPlugins])],
  };

  return Object.fromEntries(
    Object.entries(generationOptions).filter(([key, value]) => key !== 'id' && value !== undefined),
  ) as TypeDocGenerationOptions;
}

export async function generateTypeDoc(
  options: TypeDocGenerationOptions,
  dependencies: Partial<StrictTypeDocDependencies> = {},
): Promise<void> {
  const runtime = { ...defaultDependencies, ...dependencies };
  if (typeof options.out !== 'string') {
    throw new TypeError('strict TypeDoc generation requires an output path');
  }
  const approvedOutput = resolve(runtime.projectRoot, 'website/docs/api');
  const resolvedOutput = resolve(runtime.projectRoot, options.out);
  const hasTraversal = options.out.split(/[\\/]/u).includes('..');
  if (hasTraversal || resolvedOutput !== approvedOutput) {
    throw new Error(`Unsafe TypeDoc output path: ${resolvedOutput}`);
  }

  await runtime.clearOutput(resolvedOutput);
  const app = await runtime.bootstrap(options);

  if (app.logger.hasErrors()) throw new Error(describeFailure('bootstrap', app));
  if (app.options.getValue('treatWarningsAsErrors') && app.logger.hasWarnings()) {
    throw new Error(describeFailure('bootstrap', app));
  }

  const project = await app.convert();
  if (!project) throw new Error(describeFailure('conversion', app));
  if (app.logger.hasErrors()) throw new Error(describeFailure('conversion', app));
  if (app.options.getValue('treatWarningsAsErrors') && app.logger.hasWarnings()) {
    throw new Error(describeFailure('conversion', app));
  }

  const preValidationWarningCount = app.logger.warningCount;
  app.validate(project);
  const validationWarningCount = Math.max(
    app.logger.warningCount - preValidationWarningCount,
    app.logger.validationWarningCount,
  );

  if (app.logger.hasErrors()) throw new Error(describeFailure('validation', app));
  if (
    validationWarningCount > 0 &&
    (app.options.getValue('treatWarningsAsErrors') ||
      app.options.getValue('treatValidationWarningsAsErrors'))
  ) {
    throw new Error(
      `TypeDoc validation failed with ${app.logger.errorCount} errors and ${validationWarningCount} warnings`,
    );
  }

  const preOutputWarningCount = app.logger.warningCount;
  try {
    await app.generateOutputs(project);
    const outputWarningCount = app.logger.warningCount - preOutputWarningCount;
    if (app.logger.hasErrors()) throw new Error(describeFailure('output', app));
    if (outputWarningCount > 0 && app.options.getValue('treatWarningsAsErrors')) {
      throw new Error(describeFailure('output', app));
    }
  } catch (error) {
    await runtime.clearOutput(resolvedOutput);
    throw error;
  }
}

export default async function strictTypeDocGenerationPlugin(
  context: LoadContext,
  options: unknown,
): Promise<Plugin<unknown>> {
  if (typeof options !== 'object' || options === null || Array.isArray(options)) {
    throw new TypeError('strict TypeDoc generation requires an options object');
  }

  await generateTypeDoc(
    toTypeDocGenerationOptions(
      {
        siteDir: context.siteDir,
        siteConfig: { presets: context.siteConfig.presets },
      },
      options as Record<string, unknown>,
    ),
  );

  return { name: 'strict-typedoc-generation' };
}
