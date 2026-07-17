import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { isAbsolute, join, relative, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import ts from 'typescript';
import { beforeAll, describe, expect, it } from 'vitest';

const root = process.cwd();
const read = (path: string) => readFileSync(join(root, path), 'utf8');
const configUrl = pathToFileURL(join(root, 'website/docusaurus.config.ts')).href;
const sidebarStructureUrl = pathToFileURL(join(root, 'website/sidebar-structure.ts')).href;
const jitiPath = createRequire(import.meta.url).resolve('jiti');

const publishedPackagePaths = new Set([
  'tego-sheet',
  'tego-sheet/styles.css',
  'tego-sheet/locales/en',
  'tego-sheet/locales/zh-cn',
  'tego-sheet/locales/de',
  'tego-sheet/locales/nl',
]);

const publishedStylesDeclaration = 'website/src/types/tego-sheet-styles.d.ts';
const toPosixPath = (path: string): string => path.replace(/\\/g, '/');

interface SitePackageExport {
  readonly specifier: string;
  readonly runtimeTarget: string;
  readonly typeTarget?: string;
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

const sitePackageExports = (packageJson: unknown): readonly SitePackageExport[] => {
  if (!isRecord(packageJson) || !isRecord(packageJson.exports)) {
    throw new Error('package.json must define conditional exports');
  }

  const packageExports = packageJson.exports;
  const exportKeys = Object.keys(packageExports).filter((key) => key !== './package.json');
  const expectedKeys = [...publishedPackagePaths]
    .map((specifier) =>
      specifier === 'tego-sheet' ? '.' : `./${specifier.slice('tego-sheet/'.length)}`,
    )
    .sort();
  if (JSON.stringify([...exportKeys].sort()) !== JSON.stringify(expectedKeys)) {
    throw new Error(`site package export drift: ${exportKeys.join(', ')}`);
  }

  return exportKeys.map((exportKey) => {
    const specifier = exportKey === '.' ? 'tego-sheet' : `tego-sheet/${exportKey.slice(2)}`;
    const conditions = packageExports[exportKey];
    if (typeof conditions === 'string') {
      return { specifier, runtimeTarget: conditions };
    }
    if (!isRecord(conditions) || typeof conditions.import !== 'string') {
      throw new Error(`${exportKey} must define an import target`);
    }
    if (typeof conditions.types !== 'string') {
      throw new Error(`${exportKey} must define a types target`);
    }
    return {
      specifier,
      runtimeTarget: conditions.import,
      typeTarget: conditions.types,
    };
  });
};

const websiteProgramDiagnostics = (probeSource: string): readonly ts.Diagnostic[] => {
  const websiteRoot = join(root, 'website');
  const configPath = join(websiteRoot, 'tsconfig.json');
  const configFile = ts.readConfigFile(configPath, ts.sys.readFile);
  if (configFile.error) return [configFile.error];
  const parsed = ts.parseJsonConfigFileContent(
    configFile.config,
    ts.sys,
    websiteRoot,
    {},
    configPath,
  );
  if (parsed.errors.length > 0) return parsed.errors;

  const fixtureRoot = mkdtempSync(join(tmpdir(), 'tego-sheet-website-program-'));
  try {
    const probePath = join(fixtureRoot, 'public-package-probe.ts');
    writeFileSync(probePath, probeSource);
    const program = ts.createProgram({
      rootNames: [...parsed.fileNames, probePath],
      options: parsed.options,
    });
    return ts.getPreEmitDiagnostics(program);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
};

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

const dynamicModuleSpecifier = '<dynamic module specifier>';
const approvedNonCheckableFences = new Set<string>();

interface CheckableFence {
  readonly documentPath: string;
  readonly fenceNumber: number;
  readonly language: 'ts' | 'tsx';
  readonly source: string;
}

const literalModuleSpecifier = (node: ts.Node | undefined): string | null => {
  if (node && (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node))) {
    return node.text;
  }
  return null;
};

const moduleSpecifiers = (source: string): string[] => {
  const sourceFile = ts.createSourceFile(
    'documentation-example.tsx',
    source,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TSX,
  );
  const specifiers: string[] = [];

  const visit = (node: ts.Node): void => {
    if (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) {
      const specifier = literalModuleSpecifier(node.moduleSpecifier);
      if (specifier !== null) specifiers.push(specifier);
    } else if (
      ts.isImportEqualsDeclaration(node) &&
      ts.isExternalModuleReference(node.moduleReference)
    ) {
      specifiers.push(
        literalModuleSpecifier(node.moduleReference.expression) ?? dynamicModuleSpecifier,
      );
    } else if (ts.isCallExpression(node)) {
      const isDynamicImport = node.expression.kind === ts.SyntaxKind.ImportKeyword;
      const isRequire = ts.isIdentifier(node.expression) && node.expression.text === 'require';
      if (isDynamicImport || isRequire) {
        specifiers.push(literalModuleSpecifier(node.arguments[0]) ?? dynamicModuleSpecifier);
      }
    }
    ts.forEachChild(node, visit);
  };

  visit(sourceFile);
  return specifiers;
};

const documentationModuleSpecifiers = (content: string): string[] => {
  const fencedSources: string[] = [];
  const outsideFences = content.replace(/```[^\n]*\n([\s\S]*?)```/g, (_block, source: string) => {
    fencedSources.push(source);
    return '\n';
  });
  return [outsideFences, ...outsideFences.split('\n'), ...fencedSources]
    .flatMap(moduleSpecifiers)
    .filter((specifier, index, all) => all.indexOf(specifier) === index);
};

const unsupportedModuleSpecifiers = (specifiers: readonly string[]): string[] =>
  specifiers.filter(
    (specifier) =>
      specifier === dynamicModuleSpecifier ||
      (specifier.startsWith('tego-sheet') && !publishedPackagePaths.has(specifier)) ||
      /(?:^|\/)(?:src|core|controller|engine|private)(?:\/|$)/.test(specifier),
  );

const unsupportedDocumentationModules = (content: string): string[] =>
  unsupportedModuleSpecifiers(documentationModuleSpecifiers(content));

const unsupportedSourceModules = (source: string): string[] =>
  unsupportedModuleSpecifiers(moduleSpecifiers(source));

const checkableCodeFences = (): CheckableFence[] =>
  manualDocumentation.flatMap(([documentPath]) => {
    const content = read(documentPath);
    let fenceNumber = 0;
    return [...content.matchAll(/```([^\n]*)\n([\s\S]*?)```/g)].flatMap((match) => {
      fenceNumber += 1;
      const metadata = match[1].trim().split(/\s+/).filter(Boolean);
      const language = metadata[0];
      if (language !== 'ts' && language !== 'tsx') return [];
      const fenceId = `${documentPath}#${fenceNumber}`;
      if (metadata.includes('non-checkable')) {
        if (!approvedNonCheckableFences.has(fenceId)) {
          throw new Error(`${fenceId} uses an unapproved non-checkable marker`);
        }
        return [];
      }
      return [{ documentPath, fenceNumber, language, source: match[2] }];
    });
  });

const typecheckDocumentationFences = (): string[] => {
  const configFile = ts.readConfigFile(join(root, 'tsconfig.json'), ts.sys.readFile);
  if (configFile.error) {
    throw new Error(ts.flattenDiagnosticMessageText(configFile.error.messageText, '\n'));
  }
  const parsedConfig = ts.parseJsonConfigFileContent(configFile.config, ts.sys, root);
  if (parsedConfig.errors.length > 0) {
    throw new Error(
      ts.formatDiagnosticsWithColorAndContext(parsedConfig.errors, {
        getCanonicalFileName: (name) => name,
        getCurrentDirectory: () => root,
        getNewLine: () => '\n',
      }),
    );
  }

  const fixtureRoot = mkdtempSync(join(tmpdir(), 'tego-sheet-documentation-fences-'));
  try {
    const fences = checkableCodeFences();
    const fileOwners = new Map<string, CheckableFence>();
    const rootNames = fences.map((fence, index) => {
      const fileName = join(fixtureRoot, `fence-${index + 1}.${fence.language}`);
      writeFileSync(fileName, fence.source);
      fileOwners.set(fileName, fence);
      return fileName;
    });
    const localStylesDeclaration = join(fixtureRoot, 'sheet-panel.d.css.ts');
    writeFileSync(localStylesDeclaration, 'export {};\n');
    symlinkSync(
      join(root, 'node_modules'),
      join(fixtureRoot, 'node_modules'),
      process.platform === 'win32' ? 'junction' : 'dir',
    );
    rootNames.push(
      join(root, publishedStylesDeclaration),
      localStylesDeclaration,
      join(root, 'src/styles.d.ts'),
    );

    const program = ts.createProgram({
      rootNames,
      options: {
        ...parsedConfig.options,
        allowArbitraryExtensions: true,
        baseUrl: root,
        composite: false,
        incremental: false,
        noEmit: true,
      },
    });
    return rootNames.flatMap((fileName) => {
      const owner = fileOwners.get(fileName);
      if (!owner) return [];
      const sourceFile = program.getSourceFile(fileName);
      if (!sourceFile) return [`${owner.documentPath} fence ${owner.fenceNumber}: source missing`];
      return [
        ...program.getSyntacticDiagnostics(sourceFile),
        ...program.getSemanticDiagnostics(sourceFile),
      ].map((diagnostic) => {
        const position =
          diagnostic.start === undefined
            ? ''
            : (() => {
                const location = sourceFile.getLineAndCharacterOfPosition(diagnostic.start);
                return `:${location.line + 1}:${location.character + 1}`;
              })();
        return `${owner.documentPath} fence ${owner.fenceNumber}${position} TS${diagnostic.code}: ${ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n')}`;
      });
    });
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
};

const transpileModule = (source: string, fileName: string): string => {
  const result = ts.transpileModule(source, {
    fileName,
    reportDiagnostics: true,
    compilerOptions: {
      module: ts.ModuleKind.ES2022,
      target: ts.ScriptTarget.ES2022,
      verbatimModuleSyntax: true,
    },
  });
  if (result.diagnostics && result.diagnostics.length > 0) {
    throw new Error(
      ts.formatDiagnosticsWithColorAndContext(result.diagnostics, {
        getCanonicalFileName: (name) => name,
        getCurrentDirectory: () => root,
        getNewLine: () => '\n',
      }),
    );
  }
  return result.outputText;
};

const loadRealSidebarsInIsolatedProcess = (
  fixtureSource = 'module.exports = [{ type: "doc", id: "api/generated-entry" }];\n',
): unknown => {
  const fixtureRoot = mkdtempSync(join(tmpdir(), 'tego-sheet-sidebar-contract-'));
  try {
    const sourceSidebars = join(root, 'website/sidebars.ts');
    const sourceStructure = join(root, 'website/sidebar-structure.ts');
    const copiedSidebars = join(fixtureRoot, 'sidebars.ts');
    const copiedStructure = join(fixtureRoot, 'sidebar-structure.ts');
    copyFileSync(sourceSidebars, copiedSidebars);
    copyFileSync(sourceStructure, copiedStructure);
    mkdirSync(join(fixtureRoot, 'docs/api'), { recursive: true });
    writeFileSync(join(fixtureRoot, 'docs/api/typedoc-sidebar.cjs'), fixtureSource);
    writeFileSync(
      join(fixtureRoot, 'sidebars.mjs'),
      transpileModule(readFileSync(copiedSidebars, 'utf8'), copiedSidebars),
    );
    writeFileSync(
      join(fixtureRoot, 'sidebar-structure'),
      transpileModule(readFileSync(copiedStructure, 'utf8'), copiedStructure),
    );

    const moduleUrl = pathToFileURL(join(fixtureRoot, 'sidebars.mjs')).href;
    const child = spawnSync(
      process.execPath,
      [
        '--input-type=module',
        '--eval',
        `import sidebars from ${JSON.stringify(moduleUrl)}; process.stdout.write(JSON.stringify(sidebars));`,
      ],
      { encoding: 'utf8' },
    );
    if (child.status !== 0) {
      throw new Error(`Isolated sidebar load failed:\n${child.stderr || child.stdout}`);
    }
    return JSON.parse(child.stdout) as unknown;
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
};

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
  beforeAll(() => {
    const build = spawnSync('npm', ['run', 'build'], { cwd: root, encoding: 'utf8' });
    if (build.status !== 0) {
      throw new Error(`Public package build failed:\n${build.stdout}\n${build.stderr}`);
    }
  });

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

  it('keeps Quick Start linkable and isolates the interactive playground behind BrowserOnly', () => {
    const quickStart = read('website/docs/getting-started/quick-start.mdx');
    const playground = read('website/src/pages/playground.tsx');

    expect(quickStart).toContain('# Quick Start');
    expect(quickStart).toContain('./installation');
    expect(playground).toContain("import BrowserOnly from '@docusaurus/BrowserOnly'");
    expect(playground).toContain('fallback={<PlaygroundLoadingState />}');
    expect(playground).toContain('{() => <Playground />}');
    expect(playground).not.toContain('TegoSheet');
  });

  it('presents the approved home-page learning path through the public package', () => {
    const home = read('website/src/pages/index.tsx');
    const preview = read('website/src/components/homepage-preview.tsx');

    expect(home).toContain('/docs/getting-started/quick-start');
    expect(home).toContain('/playground');
    expect(home).toContain("from 'tego-sheet'");
    expect(home).toContain("import 'tego-sheet/styles.css'");
    expect(home).toContain('HomepagePreview');
    expect(preview).toContain("import BrowserOnly from '@docusaurus/BrowserOnly'");
    expect(preview).toContain("from 'tego-sheet'");
    expect(preview).toContain("import 'tego-sheet/styles.css'");
    expect(unsupportedSourceModules(preview)).toEqual([]);
  });

  it('keeps every Playground mode and inspector control at least 44 pixels tall', () => {
    const playgroundStyles = read('website/src/components/playground/playground.module.css');

    expect(playgroundStyles).toMatch(
      /\.modePicker span,[\s\S]*?\.inspector button,[\s\S]*?\.customChrome button\s*{[^}]*min-height:\s*2\.75rem/,
    );
    expect(playgroundStyles).toMatch(/\.field select\s*{[^}]*min-height:\s*2\.75rem/);
    expect(playgroundStyles).not.toMatch(/min-height:\s*2\.35rem/);
  });

  it('documents the exact React peer dependency range from package metadata', () => {
    const packageJson = JSON.parse(read('package.json')) as {
      peerDependencies?: Record<string, unknown>;
    };
    const peers = packageJson.peerDependencies;
    const quickStart = read('website/docs/getting-started/quick-start.mdx');
    const normalizedQuickStart = quickStart.replace(/\s+/g, ' ').trim();

    expect(peers).toMatchObject({ react: '^19.2.7', 'react-dom': '^19.2.7' });
    expect(normalizedQuickStart).toContain(
      `React and React DOM must both satisfy \`${peers?.react}\` (19.2.7 through 19.x)`,
    );
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
    ["import Workbook from 'tego-sheet/src/private'", 'tego-sheet/src/private'],
    ["import('../../src/core')", '../../src/core'],
    ["require('tego-sheet/controller/private')", 'tego-sheet/controller/private'],
    ["export { Workbook } from 'tego-sheet/core'", 'tego-sheet/core'],
    ["import privateModule = require('tego-sheet/private')", 'tego-sheet/private'],
    ['import(`tego-sheet/core`)', 'tego-sheet/core'],
    ['require(`tego-sheet/controller/private`)', 'tego-sheet/controller/private'],
    ['import(`tego-sheet/${section}`)', dynamicModuleSpecifier],
    ['require(`tego-sheet/${section}`)', dynamicModuleSpecifier],
  ])('rejects private module loading in documentation: %s', (source, expected) => {
    expect(unsupportedSourceModules(source)).toEqual([expected]);
  });

  it.each([
    ['import(`tego-sheet`)', 'tego-sheet'],
    ['require(`tego-sheet/locales/en`)', 'tego-sheet/locales/en'],
  ])('allows published template-literal module loading: %s', (source, expected) => {
    expect(moduleSpecifiers(source)).toContain(expected);
    expect(unsupportedSourceModules(source)).toEqual([]);
  });

  it('semantically typechecks every Task 4 TypeScript code fence against the public package', () => {
    expect(approvedNonCheckableFences).toEqual(new Set());
    expect(checkableCodeFences()).toHaveLength(13);
    expect(typecheckDocumentationFences()).toEqual([]);
  });

  it('defines the complete sidebar structure from explicit document IDs', async () => {
    const typedocSidebar = [{ type: 'doc', id: 'api/generated-entry' }];
    const sidebars = loadRealSidebarsInIsolatedProcess();

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

  it.each([
    ['module.exports = {};\n', /Generated TypeDoc sidebar export must be an array/],
    [
      'module.exports = [{ type: "doc" }];\n',
      /Generated TypeDoc sidebar export\[0\]\.id must be a non-empty string/,
    ],
    [
      'module.exports = [{ type: "doc", id: "api/example", label: 42 }];\n',
      /Generated TypeDoc sidebar export\[0\]\.label must be a non-empty string/,
    ],
    [
      'module.exports = [{ type: "category", label: "Example", collapsed: "yes", items: [] }];\n',
      /Generated TypeDoc sidebar export\[0\]\.collapsed must be a boolean/,
    ],
    [
      'module.exports = [{ type: "link", label: "Example", href: "/example", customProps: "bad" }];\n',
      /Generated TypeDoc sidebar export\[0\]\.customProps must be an object/,
    ],
  ])('rejects malformed generated sidebar exports', (fixtureSource, expected) => {
    expect(() => loadRealSidebarsInIsolatedProcess(fixtureSource)).toThrow(expected);
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

  it('typechecks every published package import in the real website program', () => {
    const diagnostics = websiteProgramDiagnostics(`
      import { TegoSheet, type LocaleDefinition, type WorkbookData } from 'tego-sheet';
      import 'tego-sheet/styles.css';
      import { en } from 'tego-sheet/locales/en';
      import { zhCN } from 'tego-sheet/locales/zh-cn';
      import { de } from 'tego-sheet/locales/de';
      import { nl } from 'tego-sheet/locales/nl';

      const workbook: WorkbookData = [];
      const locales: readonly LocaleDefinition[] = [en, zhCN, de, nl];
      void [TegoSheet, workbook, locales];
    `);

    expect(
      diagnostics.map((diagnostic) =>
        ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'),
      ),
    ).toEqual([]);
  });

  it('keeps the recursive playground module graph on published package boundaries', () => {
    const componentRoot = resolve(root, 'website/src/components/playground');
    const entryPath = resolve(componentRoot, 'playground.tsx');
    const entrySpecifiers = moduleSpecifiers(
      read('website/src/components/playground/playground.tsx'),
    );
    const visited = new Set<string>();
    const externalSpecifiers = new Set<string>();

    const visitModule = (path: string): void => {
      if (visited.has(path)) return;
      visited.add(path);
      for (const specifier of moduleSpecifiers(read(relative(root, path)))) {
        if (!specifier.startsWith('.')) {
          externalSpecifiers.add(specifier);
          continue;
        }
        if (specifier.endsWith('.css')) continue;
        const unresolved = resolve(path, '..', specifier);
        const candidate = [unresolved, `${unresolved}.ts`, `${unresolved}.tsx`].find(existsSync);
        expect(candidate, `${specifier} from ${relative(root, path)} must resolve`).toBeDefined();
        expect(toPosixPath(relative(componentRoot, candidate!))).not.toMatch(/^\.\./);
        visitModule(candidate!);
      }
    };

    visitModule(entryPath);

    expect(entrySpecifiers).toEqual(
      expect.arrayContaining([
        'tego-sheet',
        'tego-sheet/locales/en',
        'tego-sheet/locales/zh-cn',
        'tego-sheet/locales/de',
        'tego-sheet/locales/nl',
        'tego-sheet/styles.css',
      ]),
    );
    const discoveredPackageImports = [...externalSpecifiers].filter((specifier) =>
      specifier.startsWith('tego-sheet'),
    );
    expect(
      discoveredPackageImports.every((specifier) => publishedPackagePaths.has(specifier)),
    ).toBe(true);
    expect(
      [...externalSpecifiers].filter(
        (specifier) =>
          specifier !== 'react' &&
          !specifier.startsWith('@docusaurus/') &&
          !specifier.startsWith('tego-sheet'),
      ),
    ).toEqual([]);
    expect(
      [...externalSpecifiers].some((specifier) =>
        /^tego-sheet\/(?:src|private|core|controller)(?:\/|$)/.test(specifier),
      ),
    ).toBe(false);
    expect([...visited].map((path) => relative(root, path))).toEqual(
      expect.arrayContaining([
        'website/src/components/playground/playground-error-boundary.tsx',
        'website/src/components/playground/playground-fixtures.ts',
        'website/src/components/playground/playground-model.ts',
        'website/src/components/playground/playground.tsx',
      ]),
    );
  });

  it('keeps the SSR loading skeleton static and motion-free', () => {
    const styles = read('website/src/pages/playground.module.css');

    expect(styles).not.toMatch(/gradient|animation|@keyframes|prefers-reduced-motion/);
  });

  it('resolves site imports through synchronized public dist aliases', async () => {
    expect(JSON.parse(read('website/package.json'))).toEqual({ private: true });

    const packageJson = JSON.parse(read('package.json')) as unknown;
    const packageEntries = sitePackageExports(packageJson);
    const distRoot = resolve(root, 'dist');
    const resolveTarget = (target: string): string => resolve(root, target);
    const expectedTypePaths = Object.fromEntries(
      packageEntries.flatMap(({ specifier, typeTarget }) =>
        typeTarget === undefined
          ? []
          : [
              [
                specifier,
                [toPosixPath(relative(join(root, 'website'), resolveTarget(typeTarget)))],
              ],
            ],
      ),
    );
    const expectedRuntimeAliases = Object.fromEntries(
      packageEntries.map(({ specifier, runtimeTarget }) => [
        `${specifier}$`,
        resolveTarget(runtimeTarget),
      ]),
    );
    const docsTypeScript = JSON.parse(read('website/tsconfig.json')) as {
      compilerOptions?: { paths?: unknown };
    };
    const configModule = (await import(configUrl)) as {
      default: DocumentationConfig;
    };
    const aliasPlugin = configModule.default.plugins?.find(
      (
        plugin,
      ): plugin is () => {
        configureWebpack: () => { resolve: { alias: unknown } };
      } => typeof plugin === 'function' && plugin.name === 'publicPackageExportsPlugin',
    );

    expect(docsTypeScript.compilerOptions?.paths).toEqual({
      '@site/*': ['./*'],
      ...expectedTypePaths,
    });
    expect(aliasPlugin).toBeDefined();
    expect(aliasPlugin?.().configureWebpack().resolve.alias).toEqual(expectedRuntimeAliases);
    expect(packageEntries.map(({ specifier }) => specifier).sort()).toEqual(
      [...publishedPackagePaths].sort(),
    );

    for (const entry of packageEntries) {
      for (const target of [entry.runtimeTarget, entry.typeTarget].filter(
        (value): value is string => value !== undefined,
      )) {
        const absoluteTarget = resolveTarget(target);
        const distRelative = relative(distRoot, absoluteTarget);
        expect(isAbsolute(distRelative) || distRelative.startsWith('..')).toBe(false);
        expect(existsSync(absoluteTarget), `${entry.specifier} target missing: ${target}`).toBe(
          true,
        );
      }
    }

    expect(read(publishedStylesDeclaration)).toBe("declare module 'tego-sheet/styles.css';\n");
  });

  it('anchors public runtime aliases to the config instead of the process cwd', () => {
    const script = `
      const { createJiti } = require(${JSON.stringify(jitiPath)});
      const configPath = ${JSON.stringify(join(root, 'website/docusaurus.config.ts'))};
      const jiti = createJiti(configPath, { interopDefault: true });
      Promise.resolve(jiti.import(configPath, { default: true })).then((config) => {
        const plugin = config.plugins.find(
          (candidate) => typeof candidate === 'function' && candidate.name === 'publicPackageExportsPlugin',
        );
        if (!plugin) throw new Error('public package exports plugin missing');
        process.stdout.write(JSON.stringify(plugin().configureWebpack().resolve.alias));
      });
    `;
    const loaded = spawnSync(process.execPath, ['-e', script], {
      cwd: join(root, 'website'),
      encoding: 'utf8',
    });

    expect(loaded.status, loaded.stderr).toBe(0);
    const aliases = JSON.parse(loaded.stdout) as Record<string, string>;
    expect(Object.values(aliases).every((target) => target.startsWith(join(root, 'dist')))).toBe(
      true,
    );
    expect(Object.values(aliases).every(existsSync)).toBe(true);
  });

  it('rejects site package export drift and missing conditional targets', () => {
    const packageJson = JSON.parse(read('package.json')) as Record<string, unknown>;
    if (!isRecord(packageJson.exports)) throw new Error('package exports must be an object');

    const drifted = structuredClone(packageJson);
    if (!isRecord(drifted.exports)) throw new Error('package exports must be an object');
    drifted.exports['./private'] = { types: './dist/private.d.ts', import: './dist/private.js' };

    const missingTarget = structuredClone(packageJson);
    if (!isRecord(missingTarget.exports)) throw new Error('package exports must be an object');
    missingTarget.exports['./locales/en'] = { types: './dist/locales/en.d.ts' };

    expect(() => sitePackageExports(drifted)).toThrow(/site package export drift/);
    expect(() => sitePackageExports(missingTarget)).toThrow(/must define an import target/);
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
