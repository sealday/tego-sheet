import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmdirSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const read = (path: string) => readFileSync(join(root, path), 'utf8');
const configUrl = pathToFileURL(join(root, 'website/docusaurus.config.ts')).href;
const sidebarsUrl = pathToFileURL(join(root, 'website/sidebars.ts')).href;
const sidebarStructureUrl = pathToFileURL(join(root, 'website/sidebar-structure.ts')).href;

const publishedPackagePaths = new Set([
  'tego-sheet',
  'tego-sheet/styles.css',
  'tego-sheet/locales/en',
  'tego-sheet/locales/zh-cn',
  'tego-sheet/locales/de',
  'tego-sheet/locales/nl',
]);

const manualDocumentation = [
  ['website/docs/getting-started/quick-start.mdx', ['TegoSheet', 'WorkbookData']],
  ['website/docs/getting-started/styling-and-sizing.md', ['TegoSheet']],
  ['website/docs/concepts/controlled-and-uncontrolled.mdx', ['TegoSheet', 'WorkbookData']],
  ['website/docs/concepts/workbook-data.md', ['WorkbookData', 'WorkbookInput']],
  ['website/docs/concepts/refs-and-commands.mdx', ['TegoSheetHandle']],
  ['website/docs/concepts/callbacks-and-errors.mdx', ['WorkbookChange', 'TegoSheetException']],
  ['website/docs/guides/custom-chrome.mdx', ['ToolbarRenderProps', 'SheetTabsRenderProps']],
  ['website/docs/guides/locales.mdx', ['LocaleDefinition']],
  ['website/docs/guides/validation-and-filtering.md', ['ValidationData', 'AutoFilterData']],
  ['website/docs/guides/frozen-panes-and-layout.md', ['SheetData', 'SheetOptions']],
  ['website/docs/guides/printing.md', ['TegoSheetHandle']],
  ['website/docs/migration/from-x-data-spreadsheet.md', ['TegoSheet', 'TegoSheetHandle']],
] as const;

const approvedHandWrittenDocumentation = [
  'website/docs/getting-started/installation.md',
  ...manualDocumentation.map(([path]) => path),
].sort();

const collectHandWrittenDocumentation = (directory = 'website/docs'): string[] =>
  readdirSync(join(root, directory), { withFileTypes: true })
    .flatMap((entry) => {
      const path = `${directory}/${entry.name}`;
      if (entry.isDirectory())
        return path === 'website/docs/api' ? [] : collectHandWrittenDocumentation(path);
      return entry.isFile() && /\.mdx?$/.test(entry.name) ? [path] : [];
    })
    .sort();

const moduleSpecifiers = (content: string): string[] => {
  const staticModules = [
    ...content.matchAll(/\b(?:import|export)\s+(?:type\s+)?(?:[^;]*?\s+from\s+)?['"]([^'"]+)['"]/g),
  ].map((match) => match[1]);
  const runtimeModules = [
    ...content.matchAll(/\b(?:import|require)\s*\(\s*['"]([^'"]+)['"]\s*\)/g),
  ].map((match) => match[1]);

  return [...staticModules, ...runtimeModules];
};

const unsupportedDocumentationModules = (content: string): string[] =>
  moduleSpecifiers(content).filter(
    (specifier) =>
      (specifier.startsWith('tego-sheet') && !publishedPackagePaths.has(specifier)) ||
      /(?:^|\/)(?:src|core|controller|engine|private)(?:\/|$)/.test(specifier),
  );

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

  it('keeps the exact approved hand-written documentation inventory on public APIs', () => {
    expect(manualDocumentation).toHaveLength(12);
    expect(new Set(manualDocumentation.map(([path]) => path)).size).toBe(12);

    const missingPages = manualDocumentation
      .map(([path]) => path)
      .filter((path) => !existsSync(join(root, path)));

    expect(missingPages).toEqual([]);
    expect(collectHandWrittenDocumentation()).toEqual(approvedHandWrittenDocumentation);

    for (const [path, publicIdentifiers] of manualDocumentation) {
      const content = read(path);
      const headings = content.match(/^# (?!#).+$/gm) ?? [];

      expect(headings, `${path} must contain exactly one H1`).toHaveLength(1);
      for (const identifier of publicIdentifiers) {
        expect(content, `${path} must reference public API ${identifier}`).toContain(identifier);
      }
      expect(
        unsupportedDocumentationModules(content),
        `${path} must use only published modules`,
      ).toEqual([]);
    }
  });

  it.each([
    ["import('../../src/core')", '../../src/core'],
    ["require('tego-sheet/controller/private')", 'tego-sheet/controller/private'],
    ["export { Workbook } from 'tego-sheet/core'", 'tego-sheet/core'],
  ])('rejects private module loading in documentation: %s', (source, expected) => {
    expect(unsupportedDocumentationModules(source)).toEqual([expected]);
  });

  it('defines the complete sidebar structure from explicit document IDs', async () => {
    const apiDirectory = join(root, 'website/docs/api');
    const typedocSidebarPath = join(apiDirectory, 'typedoc-sidebar.cjs');
    const apiDirectoryExisted = existsSync(apiDirectory);
    const previousTypedocSidebar = existsSync(typedocSidebarPath)
      ? readFileSync(typedocSidebarPath, 'utf8')
      : null;
    const typedocSidebar = [{ type: 'doc', id: 'api/generated-entry' }];
    mkdirSync(apiDirectory, { recursive: true });
    writeFileSync(typedocSidebarPath, `module.exports = ${JSON.stringify(typedocSidebar)};\n`);

    let sidebars: unknown;
    try {
      const sidebarsModule = (await import(sidebarsUrl)) as { default: unknown };
      sidebars = sidebarsModule.default;
    } finally {
      if (previousTypedocSidebar === null) unlinkSync(typedocSidebarPath);
      else writeFileSync(typedocSidebarPath, previousTypedocSidebar);
      if (!apiDirectoryExisted && readdirSync(apiDirectory).length === 0) rmdirSync(apiDirectory);
    }

    const { createDocumentationSidebars } = (await import(sidebarStructureUrl)) as {
      createDocumentationSidebars: (typedocSidebar: readonly unknown[]) => unknown;
    };

    expect(sidebars).toEqual({
      docsSidebar: [
        {
          type: 'category',
          label: 'Getting Started',
          items: [
            'getting-started/installation',
            'getting-started/quick-start',
            'getting-started/styling-and-sizing',
          ],
        },
        {
          type: 'category',
          label: 'Core Concepts',
          items: [
            'concepts/controlled-and-uncontrolled',
            'concepts/workbook-data',
            'concepts/refs-and-commands',
            'concepts/callbacks-and-errors',
          ],
        },
        {
          type: 'category',
          label: 'Guides',
          items: [
            'guides/custom-chrome',
            'guides/locales',
            'guides/validation-and-filtering',
            'guides/frozen-panes-and-layout',
            'guides/printing',
          ],
        },
        {
          type: 'category',
          label: 'Migration',
          items: ['migration/from-x-data-spreadsheet'],
        },
        {
          type: 'category',
          label: 'API Reference',
          link: { type: 'doc', id: 'api/index' },
          items: typedocSidebar,
        },
      ],
    });
    const builderSidebars = createDocumentationSidebars(typedocSidebar) as {
      docsSidebar: { items: unknown }[];
    };
    expect(builderSidebars.docsSidebar.at(-1)?.items).toBe(typedocSidebar);
    expect((sidebars as { docsSidebar: { items: unknown }[] }).docsSidebar.at(-1)?.items).toEqual(
      typedocSidebar,
    );
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
