import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const read = (path: string) => readFileSync(join(root, path), 'utf8');
const configUrl = pathToFileURL(join(root, 'website/docusaurus.config.ts')).href;

interface DocumentationConfig {
  baseUrl?: unknown;
  i18n?: {
    defaultLocale?: unknown;
    locales?: unknown;
  };
  onBrokenLinks?: unknown;
  organizationName?: unknown;
  plugins?: readonly unknown[];
  presets?: readonly unknown[];
  projectName?: unknown;
  themeConfig?: {
    navbar?: {
      items?: readonly {
        href?: unknown;
        label?: unknown;
        position?: unknown;
        to?: unknown;
      }[];
    };
  };
  trailingSlash?: unknown;
  url?: unknown;
}

type PluginTuple = readonly [string, Record<string, unknown>];
type LocalPluginTuple = readonly [
  (...args: readonly unknown[]) => unknown,
  Record<string, unknown>,
];

const isPluginTuple = (value: unknown): value is PluginTuple =>
  Array.isArray(value) &&
  typeof value[0] === 'string' &&
  typeof value[1] === 'object' &&
  value[1] !== null &&
  !Array.isArray(value[1]);

const isLocalPluginTuple = (value: unknown): value is LocalPluginTuple =>
  Array.isArray(value) &&
  typeof value[0] === 'function' &&
  typeof value[1] === 'object' &&
  value[1] !== null &&
  !Array.isArray(value[1]);

const loadConfig = async (): Promise<DocumentationConfig> => {
  const configModule = (await import(configUrl)) as { default: DocumentationConfig };
  return configModule.default;
};

describe('documentation site contract', () => {
  it('builds for the GitHub Pages project path', async () => {
    const config = await loadConfig();

    expect(config).toMatchObject({
      url: 'https://sealday.github.io',
      baseUrl: '/tego-sheet/',
      organizationName: 'sealday',
      projectName: 'tego-sheet',
      trailingSlash: false,
      onBrokenLinks: 'throw',
    });
  });

  it('generates API docs only from the package public entry point', async () => {
    const config = await loadConfig();
    const strictGenerator = config.plugins?.find(
      (plugin): plugin is LocalPluginTuple =>
        isLocalPluginTuple(plugin) && plugin[0].name === 'strictTypeDocGenerationPlugin',
    );
    const eagerTypeDocPlugins = config.plugins?.filter(
      (plugin) => isPluginTuple(plugin) && plugin[0] === 'docusaurus-plugin-typedoc',
    );

    expect(strictGenerator).toBeDefined();
    if (!strictGenerator) throw new Error('strict TypeDoc generation must be configured');
    expect(eagerTypeDocPlugins).toEqual([]);
    expect(config.plugins?.filter(isLocalPluginTuple)).toEqual([strictGenerator]);
    const typedocOptions = strictGenerator[1];
    expect(typedocOptions.entryPoints).toEqual(['src/index.ts']);
    expect(typedocOptions.entryPoints).not.toContain('src/core/index.ts');
    expect(typedocOptions.out).toBe('website/docs/api');
    expect(typedocOptions.treatWarningsAsErrors).toBe(true);
    expect(typedocOptions.treatValidationWarningsAsErrors).toBe(true);
  });

  it('publishes one English site with the approved global navigation', async () => {
    const config = await loadConfig();

    expect(config.i18n).toEqual({ defaultLocale: 'en', locales: ['en'] });
    expect(config.themeConfig?.navbar?.items).toEqual([
      {
        to: '/docs/getting-started/installation',
        label: 'Docs',
        position: 'left',
      },
      { to: '/docs/api', label: 'API', position: 'left' },
      { to: '/playground', label: 'Playground', position: 'left' },
      {
        href: 'https://github.com/sealday/tego-sheet',
        label: 'GitHub',
        position: 'right',
      },
    ]);
  });

  it('keeps approved future routes linkable without implementing later features', () => {
    const quickStart = read('website/docs/getting-started/quick-start.mdx');
    const playground = read('website/src/pages/playground.tsx');

    expect(quickStart).toContain('# Quick Start');
    expect(quickStart).toContain('./installation');
    expect(playground).toContain('<h1>Playground</h1>');
    expect(playground).not.toContain('TegoSheet');
  });

  it('disables the classic preset blog structurally', async () => {
    const config = await loadConfig();
    const classicPreset = config.presets?.find(
      (preset): preset is PluginTuple =>
        isPluginTuple(preset) && preset[0] === '@docusaurus/preset-classic',
    );

    expect(classicPreset).toBeDefined();
    if (!classicPreset) throw new Error('@docusaurus/preset-classic must be configured');
    expect(classicPreset[1].blog).toBe(false);
  });

  it('keeps generated documentation output untracked', () => {
    const ignore = read('.gitignore');

    expect(ignore).toContain('/website/docs/api/');
    expect(ignore).toContain('/website/build/');
    expect(ignore).toContain('/website/.docusaurus/');
  });

  it('typechecks the CommonJS TypeDoc runtime bridge as checked JavaScript', () => {
    const docsTypeScript = JSON.parse(read('website/tsconfig.json')) as {
      compilerOptions?: { allowJs?: unknown; checkJs?: unknown };
      include?: unknown[];
    };

    expect(docsTypeScript.compilerOptions).toMatchObject({ allowJs: true, checkJs: true });
    expect(docsTypeScript.include).toContain('plugins/**/*.cjs');
  });
});
