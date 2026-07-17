import { resolve } from 'node:path';
import type { Options, ThemeConfig } from '@docusaurus/preset-classic';
import type { Config } from '@docusaurus/types';
import type { PluginOptions as DocusaurusTypeDocOptions } from 'docusaurus-plugin-typedoc';
import { themes as prismThemes } from 'prism-react-renderer';
import type { TypeDocOptions } from 'typedoc';
import strictTypeDocGenerationPlugin from './plugins/strict-typedoc-generation';

const PUBLIC_PACKAGE_RUNTIME_ALIASES = Object.freeze({
  'tego-sheet$': resolve(process.cwd(), 'dist/tego-sheet.js'),
  'tego-sheet/styles.css$': resolve(process.cwd(), 'dist/styles.css'),
  'tego-sheet/locales/en$': resolve(process.cwd(), 'dist/locales/en.js'),
  'tego-sheet/locales/zh-cn$': resolve(process.cwd(), 'dist/locales/zh-cn.js'),
  'tego-sheet/locales/de$': resolve(process.cwd(), 'dist/locales/de.js'),
  'tego-sheet/locales/nl$': resolve(process.cwd(), 'dist/locales/nl.js'),
});

function publicPackageExportsPlugin() {
  return {
    name: 'public-package-exports',
    configureWebpack() {
      return { resolve: { alias: PUBLIC_PACKAGE_RUNTIME_ALIASES } };
    },
  };
}

const typedocOptions = {
  entryPoints: ['src/index.ts'],
  tsconfig: 'tsconfig.json',
  out: 'website/docs/api',
  readme: 'none',
  excludePrivate: true,
  excludeProtected: true,
  excludeInternal: true,
  validation: {
    invalidLink: true,
    notDocumented: true,
    notExported: true,
  },
  treatWarningsAsErrors: true,
  treatValidationWarningsAsErrors: true,
  requiredToBeDocumented: [
    'Class',
    'Interface',
    'Function',
    'Enum',
    'TypeAlias',
    'Variable',
    'Property',
    'Method',
    'Accessor',
    'Constructor',
  ],
  sidebar: {
    autoConfiguration: true,
    deprecatedItemClassName: 'typedoc-sidebar-item-deprecated',
    pretty: true,
    typescript: false,
  },
} satisfies TypeDocOptions & DocusaurusTypeDocOptions;

const config: Config = {
  title: 'tego-sheet',
  tagline: 'A typed React spreadsheet for real application workflows',
  favicon: 'img/favicon.svg',
  url: 'https://sealday.github.io',
  baseUrl: '/tego-sheet/',
  organizationName: 'sealday',
  projectName: 'tego-sheet',
  trailingSlash: false,
  onBrokenLinks: 'throw',
  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },
  presets: [
    [
      '@docusaurus/preset-classic',
      {
        docs: {
          routeBasePath: 'docs',
          sidebarPath: './sidebars.ts',
        },
        blog: false,
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Options,
    ],
  ],
  plugins: [[strictTypeDocGenerationPlugin, typedocOptions], publicPackageExportsPlugin],
  themeConfig: {
    navbar: {
      title: 'tego-sheet',
      items: [
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
      ],
    },
    colorMode: {
      respectPrefersColorScheme: true,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
    },
  } satisfies ThemeConfig,
};

export default config;
