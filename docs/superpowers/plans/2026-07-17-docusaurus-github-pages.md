# tego-sheet Docusaurus GitHub Pages Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish an English-first Docusaurus site for `tego-sheet` with hand-written usage docs, strict TypeDoc API reference generation, five interactive public-API-only playground modes, and GitHub Pages deployment gated by the repository's complete CI suite.

**Architecture:** Keep the documentation application under `website/` while retaining one root npm package and lockfile. Docusaurus builds only after the library `dist` exists, TypeDoc reads the single public entry point `src/index.ts`, and the browser-only playground imports `tego-sheet` through package exports. GitHub Actions builds and tests the site on pull requests, then publishes the exact tested artifact only after every existing and new quality gate passes on `main`.

**Tech Stack:** Docusaurus 3.10.2, React 19.2.7, TypeScript 6.0.3, TypeDoc 0.28.20, typedoc-plugin-markdown 4.12.0, docusaurus-plugin-typedoc 1.4.2, Vitest 4.1.10, Playwright 1.61.1, Oxlint, Oxfmt, GitHub Actions, GitHub Pages.

---

## Execution rules

- Work from `/Users/seal/projects/tego-sheet` on the current branch.
- Preserve all existing package exports and runtime behavior. The site may consume public APIs, but it must not create new library APIs.
- Use test-driven development for each behavioral slice: add one failing assertion, run it and read the failure, implement the minimum, then rerun the focused test.
- Run `npm run format` before each commit. Never format generated `website/docs/api/` output or commit that directory.
- Keep generated Docusaurus output (`website/build/`, `website/.docusaurus/`) and generated TypeDoc Markdown (`website/docs/api/`) untracked.
- Every commit must pass commitlint and use the repository Lore trailers shown below.
- Do not add a live code editor, blog, versioning, third-party search, analytics, storage, login, or private engine access.

## Task 1: Lock the documentation toolchain and repository contract

**Files:**

- Modify: `package.json`
- Modify: `package-lock.json`
- Modify: `.gitignore`
- Modify: `.oxlintrc.json`
- Modify: `.oxfmtrc.json`
- Modify: `tests/package/repository-policy.test.mjs`
- Create: `tests/architecture/documentation-site-contract.test.ts`

- [ ] **Step 1: Add failing repository-policy assertions for exact dependencies and scripts**

In `tests/package/repository-policy.test.mjs`, extend the existing exact dependency map with:

```js
'@docusaurus/core': '3.10.2',
'@docusaurus/preset-classic': '3.10.2',
'@docusaurus/tsconfig': '3.10.2',
'@docusaurus/types': '3.10.2',
'@mdx-js/react': '3.1.1',
'docusaurus-plugin-typedoc': '1.4.2',
'prism-react-renderer': '2.4.1',
'typedoc': '0.28.20',
'typedoc-plugin-markdown': '4.12.0',
```

Add exact script expectations:

```js
assert.equal(scripts['docs:start'], 'npm run build && docusaurus start website');
assert.equal(scripts['docs:build'], 'npm run build && docusaurus build website');
assert.equal(scripts['docs:serve'], 'docusaurus serve website');
assert.equal(
  scripts['typecheck:docs'],
  'npm run build && tsc --noEmit --project website/tsconfig.json',
);
assert.equal(
  scripts['test:docs'],
  'playwright test --config playwright.docs.config.ts',
);
assert.equal(
  scripts['test:docs-visual'],
  'playwright test --config playwright.docs-visual.config.ts',
);
```

- [ ] **Step 2: Add a failing architecture contract for the site boundary**

Create `tests/architecture/documentation-site-contract.test.ts` with initial assertions that intentionally fail before scaffolding:

```ts
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const read = (path: string) => readFileSync(join(root, path), 'utf8');

describe('documentation site contract', () => {
  it('builds for the GitHub Pages project path', () => {
    const config = read('website/docusaurus.config.ts');

    expect(config).toContain("url: 'https://sealday.github.io'");
    expect(config).toContain("baseUrl: '/tego-sheet/'");
    expect(config).toContain("organizationName: 'sealday'");
    expect(config).toContain("projectName: 'tego-sheet'");
    expect(config).toContain('trailingSlash: false');
    expect(config).toContain("onBrokenLinks: 'throw'");
  });

  it('generates API docs only from the package public entry point', () => {
    const config = read('website/docusaurus.config.ts');

    expect(config).toContain("entryPoints: ['../src/index.ts']");
    expect(config).not.toContain('../src/core/index.ts');
    expect(config).toContain("out: 'docs/api'");
    expect(config).toContain('treatValidationWarningsAsErrors: true');
  });

  it('keeps generated documentation output untracked', () => {
    const ignore = read('.gitignore');

    expect(ignore).toContain('/website/docs/api/');
    expect(ignore).toContain('/website/build/');
    expect(ignore).toContain('/website/.docusaurus/');
  });
});
```

- [ ] **Step 3: Run the focused tests and confirm the intended failures**

Run:

```bash
node --test tests/package/repository-policy.test.mjs
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: the policy test reports missing dependency/script entries, and the architecture test reports `ENOENT` for `website/docusaurus.config.ts`.

- [ ] **Step 4: Install the exact documentation dependencies**

Run:

```bash
npm install --save-dev --save-exact @docusaurus/core@3.10.2 @docusaurus/preset-classic@3.10.2 @docusaurus/tsconfig@3.10.2 @docusaurus/types@3.10.2 @mdx-js/react@3.1.1 docusaurus-plugin-typedoc@1.4.2 prism-react-renderer@2.4.1 typedoc@0.28.20 typedoc-plugin-markdown@4.12.0
```

Confirm `package-lock.json` records only one root package and no `website/package.json` was created.

- [ ] **Step 5: Add the root documentation scripts and ignore rules**

Add the exact scripts asserted above to `package.json`. Add these exact paths to `.gitignore`:

```gitignore
/website/build/
/website/.docusaurus/
/website/docs/api/
```

Extend `.oxlintrc.json` React overrides so `website/src/**/*.{ts,tsx}` receives the same React Hooks rules as `src/` and the same React Refresh rules as other TSX entrypoints. Extend `.oxfmtrc.json` ignore patterns with:

```json
"website/docs/api/**",
"website/build/**",
"website/.docusaurus/**"
```

- [ ] **Step 6: Rerun the policy test and record the remaining expected failure**

Run:

```bash
node --test tests/package/repository-policy.test.mjs
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: repository policy passes; the documentation architecture test still fails because the Docusaurus config is deliberately not created until Task 2.

- [ ] **Step 7: Commit the toolchain boundary**

```bash
git add package.json package-lock.json .gitignore .oxlintrc.json .oxfmtrc.json tests/package/repository-policy.test.mjs tests/architecture/documentation-site-contract.test.ts
git commit -m "build: make documentation tooling reproducible" \
  -m "Constraint: the repository keeps one root npm package and lockfile" \
  -m "Rejected: a nested website package | it would duplicate dependency policy and release maintenance" \
  -m "Confidence: high" \
  -m "Scope-risk: narrow" \
  -m "Tested: repository policy passes and the intentional scaffold contract fails at the missing config"
```

## Task 2: Scaffold the Docusaurus shell and prove a strict static build

**Files:**

- Create: `website/docusaurus.config.ts`
- Create: `website/sidebars.ts`
- Create: `website/tsconfig.json`
- Create: `website/static/.nojekyll`
- Create: `website/static/img/favicon.svg`
- Create: `website/src/css/custom.css`
- Create: `website/src/pages/index.tsx`
- Create: `website/src/pages/index.module.css`
- Create: `website/docs/getting-started/installation.md`
- Modify: `tests/architecture/documentation-site-contract.test.ts`

- [ ] **Step 1: Extend the architecture test with navigation and locale assertions**

Add assertions for the approved global navigation and English-first i18n boundary:

```ts
expect(config).toContain("defaultLocale: 'en'");
expect(config).toContain("locales: ['en']");
expect(config).toContain("label: 'Docs'");
expect(config).toContain("label: 'API'");
expect(config).toContain("label: 'Playground'");
expect(config).toContain("label: 'GitHub'");
expect(config).not.toContain('blog:');
```

- [ ] **Step 2: Run the architecture test and confirm it fails on the missing config**

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: `ENOENT` for `website/docusaurus.config.ts`.

- [ ] **Step 3: Create the typed Docusaurus configuration**

Create `website/docusaurus.config.ts` with these required settings:

```ts
import type { Config } from '@docusaurus/types';
import type { Options, ThemeConfig } from '@docusaurus/preset-classic';
import { themes as prismThemes } from 'prism-react-renderer';

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
  i18n: { defaultLocale: 'en', locales: ['en'] },
  presets: [
    [
      'classic',
      {
        docs: { routeBasePath: 'docs', sidebarPath: './sidebars.ts' },
        blog: false,
        theme: { customCss: './src/css/custom.css' },
      } satisfies Options,
    ],
  ],
  plugins: [
    [
      'docusaurus-plugin-typedoc',
      {
        entryPoints: ['../src/index.ts'],
        tsconfig: '../tsconfig.json',
        out: 'docs/api',
        readme: 'none',
        excludePrivate: true,
        excludeProtected: true,
        excludeInternal: true,
        validation: { invalidLink: true, notDocumented: true, notExported: true },
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
        sidebar: { autoConfiguration: true, pretty: true, typescript: false },
      },
    ],
  ],
  themeConfig: {
    navbar: {
      title: 'tego-sheet',
      items: [
        { to: '/docs/getting-started/installation', label: 'Docs', position: 'left' },
        { to: '/docs/api', label: 'API', position: 'left' },
        { to: '/playground', label: 'Playground', position: 'left' },
        { href: 'https://github.com/sealday/tego-sheet', label: 'GitHub', position: 'right' },
      ],
    },
    colorMode: { respectPrefersColorScheme: true },
    prism: { theme: prismThemes.github, darkTheme: prismThemes.dracula },
  } satisfies ThemeConfig,
};

export default config;
```

If TypeDoc rejects one of the `requiredToBeDocumented` reflection names, inspect the installed TypeDoc option schema and replace only the invalid names with the supported equivalent; do not weaken `notDocumented` or warning-as-error behavior.

- [ ] **Step 4: Create the site shell**

Create `website/tsconfig.json`:

```json
{
  "extends": "@docusaurus/tsconfig",
  "compilerOptions": {
    "baseUrl": ".",
    "strict": true
  },
  "include": ["docusaurus.config.ts", "sidebars.ts", "src/**/*.ts", "src/**/*.tsx"]
}
```

Create `website/sidebars.ts` with hand-written categories and a generated API sidebar loaded from `./docs/api/typedoc-sidebar.cjs`. Use `createRequire(import.meta.url)` plus a narrow `SidebarItem[]` type assertion so `typecheck:docs` does not require generated output to exist before Docusaurus invokes the TypeDoc plugin. Keep API last. Create an empty `website/static/.nojekyll` so GitHub Pages cannot apply Jekyll filtering.

Create a minimal `website/docs/getting-started/installation.md` with front matter `sidebar_position: 1`, the Node engine requirement, `npm install tego-sheet`, CSS import, and peer dependency statement.

Create a minimal typed home page in `website/src/pages/index.tsx` using Docusaurus `Layout` and `Link`; include the product title and links to Quick Start and Playground. Add only the foundational brand tokens and readable light/dark surfaces to `website/src/css/custom.css` and `website/src/pages/index.module.css`; detailed home content comes in Task 7.

- [ ] **Step 5: Run typecheck and the first strict docs build**

Run:

```bash
npm run typecheck:docs
npm run docs:build
```

Expected: typecheck passes. The docs build may now fail only on missing public TSDoc; capture the exact TypeDoc diagnostics for Task 3. It must not fail due to Docusaurus configuration, missing favicon, sidebar loading, or base URL.

- [ ] **Step 6: Make the favicon reference real and rerun the architecture contract**

Create `website/static/img/favicon.svg` as a small hand-authored grid mark using the existing navy/blue palette; do not generate a binary asset. Then run:

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: all documentation architecture assertions pass.

- [ ] **Step 7: Commit the static-site shell**

```bash
git add website tests/architecture/documentation-site-contract.test.ts
git commit -m "feat(docs): establish a strict project-path site shell" \
  -m "Constraint: GitHub Pages serves this repository below /tego-sheet/" \
  -m "Rejected: root-relative site links | they break project Pages deployments" \
  -m "Confidence: high" \
  -m "Scope-risk: moderate" \
  -m "Directive: preserve the English-first i18n boundary and generated API directory" \
  -m "Tested: docs typecheck and documentation architecture contract" \
  -m "Not-tested: strict TypeDoc completion remains blocked by public API comments"
```

## Task 3: Document the public API and make TypeDoc validation pass

**Files:**

- Modify: `src/index.ts`
- Modify: `src/react/tego-sheet.tsx`
- Modify: `src/react/tego-sheet.types.ts`
- Modify: `src/ui/slot-types.ts`
- Modify: `src/core/types/json.ts`
- Modify: `src/core/types/workbook.ts`
- Modify: `src/core/types/coordinates.ts`
- Modify: `src/core/types/changes.ts`
- Modify: `src/core/types/validation.ts`
- Modify: `src/core/types/options.ts`
- Modify: `src/core/errors/tego-sheet-error.ts`
- Modify: `src/core/errors/tego-sheet-exception.ts`
- Modify: `tests/architecture/documentation-site-contract.test.ts`

- [ ] **Step 1: Turn the captured TypeDoc diagnostics into a regression assertion**

Add a test that enumerates the public source files above and verifies every root-exported declaration has a `/** ... */` block immediately associated with it. The test is a fast guard; TypeDoc remains the authoritative member/link validator.

Use a table of exact public export names sourced from `src/index.ts`, not a repository-wide comment-percentage heuristic. Include `TegoSheet`, `TegoSheetException`, props/handle/callback types, workbook/address/range/selection types, options, validation, slot and error types.

- [ ] **Step 2: Run the focused architecture test and confirm undocumented exports are listed**

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: the failure names the current public declarations missing TSDoc.

- [ ] **Step 3: Add concise consumer-facing TSDoc without changing runtime code**

Document every public declaration and public member reported by TypeDoc. Comments must explain:

- whether `TegoSheetProps.value` or `defaultValue` owns workbook state;
- the stable responsibilities of `TegoSheetHandle` methods;
- callback timing and the meaning of workbook/error/event payloads;
- sparse workbook JSON shape and coordinate conventions;
- options, validation rules, frozen panes, filtering and printing concepts;
- slot renderer inputs/actions and which defaults a renderer replaces;
- `TegoSheetError` versus thrown `TegoSheetException`;
- locale types as per-instance dictionaries.

Use `@remarks` only for lifecycle or mutability caveats, `@example` for the component/handle entrypoints, and resolvable `{@link ...}` references for related public names. Do not mention internal controllers, Canvas engine classes, or unexported implementation names.

- [ ] **Step 4: Run strict TypeDoc generation through Docusaurus**

```bash
npm run docs:build
```

Expected: TypeDoc emits `website/docs/api/`, Docusaurus builds `website/build/`, and there are zero documentation, invalid-link, not-exported, route, or broken-link warnings.

Inspect generated output:

```bash
rg -n "controller|canvas-engine|src/core/index" website/docs/api website/build
git status --short --ignored website/docs/api website/build website/.docusaurus
```

Expected: the forbidden internal names have no public reference matches; all three generated directories show as ignored and none are staged.

- [ ] **Step 5: Run library safety checks**

```bash
npm run typecheck
npm run test:unit
npm run build
```

Expected: public API documentation changes do not alter declarations, output, or behavior.

- [ ] **Step 6: Commit public API documentation**

```bash
git add src tests/architecture/documentation-site-contract.test.ts
git commit -m "docs(api): make the public contract self-describing" \
  -m "Constraint: generated reference material must come only from src/index.ts" \
  -m "Rejected: documenting internal modules | it would turn implementation details into accidental API" \
  -m "Confidence: high" \
  -m "Scope-risk: moderate" \
  -m "Directive: new public declarations must satisfy TypeDoc validation before merge" \
  -m "Tested: strict docs build, typecheck, unit suite, and library build"
```

## Task 4: Write the usage, concept, guide, and migration documentation

**Files:**

- Modify: `website/sidebars.ts`
- Create: `website/docs/getting-started/quick-start.mdx`
- Create: `website/docs/getting-started/styling-and-sizing.md`
- Create: `website/docs/concepts/controlled-and-uncontrolled.mdx`
- Create: `website/docs/concepts/workbook-data.md`
- Create: `website/docs/concepts/refs-and-commands.mdx`
- Create: `website/docs/concepts/callbacks-and-errors.mdx`
- Create: `website/docs/guides/custom-chrome.mdx`
- Create: `website/docs/guides/locales.mdx`
- Create: `website/docs/guides/validation-and-filtering.md`
- Create: `website/docs/guides/frozen-panes-and-layout.md`
- Create: `website/docs/guides/printing.md`
- Create: `website/docs/migration/from-x-data-spreadsheet.md`
- Modify: `tests/architecture/documentation-site-contract.test.ts`

- [ ] **Step 1: Add a failing documentation inventory test**

Add an exact table of the twelve hand-written pages above and assert each exists, has a single H1, contains at least one public API identifier where relevant, and contains no imports from `src/`, `core/`, `controller`, or other private paths. Also assert `website/sidebars.ts` lists the approved category order: Getting Started, Core Concepts, Guides, Migration, API Reference.

- [ ] **Step 2: Run the inventory test and confirm every missing page is reported**

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: a deterministic list of the eleven pages not created in earlier tasks.

- [ ] **Step 3: Write Getting Started and Core Concepts**

Every example must import only public package paths:

```tsx
import { TegoSheet, type TegoSheetHandle, type WorkbookData } from 'tego-sheet';
import 'tego-sheet/styles.css';
```

Cover these exact outcomes:

- Installation: npm installation, Node/React requirements, stylesheet import.
- Quick Start: a sized parent container, `defaultValue`, `onChange`, and safe initial JSON.
- Styling and Sizing: explicit height requirement, responsive container, theme-neutral CSS.
- Controlled and Uncontrolled: separate examples and a warning never to switch a mounted instance between the two modes.
- Workbook Data: sparse rows/cells, zero-based coordinates, serialization round trip.
- Refs and Commands: `useRef<TegoSheetHandle>`, focus, value retrieval, and the actual exported handle methods only.
- Callbacks and Errors: `onChange`, public event callbacks, `onError`, and exception handling.

Verify names against `src/index.ts` and the generated API; do not invent props or handle methods.

- [ ] **Step 4: Write Guides and Migration**

Cover:

- typed custom toolbar and sheet-tabs renderers using public slot props/actions;
- per-instance imports for `tego-sheet/locales/en`, `/zh-cn`, `/de`, and `/nl`;
- validation and filtering using exported option/data types;
- frozen panes, selection/layout behavior, and explicit container sizing;
- printing through the supported public handle/option behavior;
- a concise x-data-spreadsheet migration map that links to the package-distributed canonical `docs/migration-from-x-data-spreadsheet.md` on GitHub for the exhaustive compatibility notes.

Each guide ends with a link to the relevant generated API page or `/playground?mode=...` preset where one exists.

- [ ] **Step 5: Wire the manual and generated sidebars**

Use explicit document IDs for the hand-written hierarchy. Load `./docs/api/typedoc-sidebar.cjs` and nest it under:

```ts
{
  type: 'category',
  label: 'API Reference',
  link: { type: 'doc', id: 'api/index' },
  items: typedocSidebar,
}
```

Keep API last and keep the generated `.cjs` file ignored.

- [ ] **Step 6: Verify prose structure, code snippets, and links**

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
npm run docs:build
```

Expected: inventory passes and Docusaurus reports no broken links or unresolved MDX imports.

- [ ] **Step 7: Commit the hand-written documentation**

```bash
git add website/docs website/sidebars.ts tests/architecture/documentation-site-contract.test.ts
git commit -m "docs: teach the supported React integration paths" \
  -m "Constraint: examples must remain inside the published package surface" \
  -m "Rejected: copying the full migration artifact into the site | two canonical copies would drift" \
  -m "Confidence: high" \
  -m "Scope-risk: moderate" \
  -m "Directive: verify every example against src/index.ts and generated API output" \
  -m "Tested: documentation inventory and strict Docusaurus link build"
```

## Task 5: Build the pure playground model and preset registry with unit tests

**Files:**

- Create: `website/src/components/playground/playground-model.ts`
- Create: `website/src/components/playground/playground-fixtures.ts`
- Create: `tests/unit/website/playground-model.test.ts`

- [ ] **Step 1: Write failing unit tests for modes, fixtures, and event capping**

Create tests for:

```ts
expect(parsePlaygroundMode(null)).toBe('uncontrolled');
expect(parsePlaygroundMode('controlled')).toBe('controlled');
expect(parsePlaygroundMode('private-engine')).toBe('uncontrolled');
expect(PLAYGROUND_MODES).toEqual([
  'uncontrolled',
  'controlled',
  'custom-chrome',
  'locales',
  'legacy-json',
]);
expect(appendPlaygroundEvent(fiftyEvents, nextEvent)).toHaveLength(50);
expect(createFixture('legacy-json')).not.toBe(createFixture('legacy-json'));
```

Also test that fixture mutation cannot leak between calls and that every registry item has a label, description, docs link, and public API list.

- [ ] **Step 2: Run the focused test and confirm the module is missing**

```bash
npx vitest run --project unit tests/unit/website/playground-model.test.ts
```

Expected: import failure for `playground-model.ts`.

- [ ] **Step 3: Implement the pure model**

Export:

```ts
export const PLAYGROUND_MODES = [
  'uncontrolled',
  'controlled',
  'custom-chrome',
  'locales',
  'legacy-json',
] as const;

export type PlaygroundMode = (typeof PLAYGROUND_MODES)[number];

export function parsePlaygroundMode(value: string | null): PlaygroundMode;
export function appendPlaygroundEvent(
  events: readonly PlaygroundEvent[],
  event: PlaygroundEvent,
): readonly PlaygroundEvent[];
```

`appendPlaygroundEvent` must retain the newest 50 immutable records. Event records contain only a stable sequence number, public callback name, and JSON-serializable payload; do not store DOM events or component internals.

- [ ] **Step 4: Implement fixture factories and preset metadata**

Create independent workbook factory functions for all five modes. Preset labels are exactly `Uncontrolled`, `Controlled`, `Custom Chrome`, `Locales`, and `Legacy JSON`; their URL keys remain the lower-case values in `PLAYGROUND_MODES`. The legacy fixture must use the existing compatible sparse JSON shape from library tests. Locale metadata must map only the four published locale subpaths. Keep React renderers out of this pure file so unit tests remain DOM-free.

- [ ] **Step 5: Run focused and full unit checks**

```bash
npx vitest run --project unit tests/unit/website/playground-model.test.ts
npm run test:unit
```

Expected: all model and existing unit/component/architecture tests pass.

- [ ] **Step 6: Commit the model**

```bash
git add website/src/components/playground/playground-model.ts website/src/components/playground/playground-fixtures.ts tests/unit/website/playground-model.test.ts
git commit -m "feat(playground): define isolated public-api presets" \
  -m "Constraint: preset state cannot leak across mode switches or resets" \
  -m "Rejected: one mutable workbook shared by every mode | it obscures controlled-state behavior" \
  -m "Confidence: high" \
  -m "Scope-risk: narrow" \
  -m "Directive: keep model helpers DOM-free and cap public callback history at 50" \
  -m "Tested: focused playground model tests and full unit suite"
```

## Task 6: Implement the interactive playground with component tests

**Files:**

- Create: `website/src/components/playground/playground.tsx`
- Create: `website/src/components/playground/playground-error-boundary.tsx`
- Create: `website/src/components/playground/playground.module.css`
- Create: `website/src/pages/playground.tsx`
- Create: `website/src/pages/playground.module.css`
- Create: `tests/component/docs-playground.test.tsx`

- [ ] **Step 1: Write a public-package mock and failing interaction tests**

Mock only the `tego-sheet` package boundary, not internal modules. The mock should expose a test double that records `value`, `defaultValue`, locale and slot props, invokes public callbacks, and exposes only the actual `TegoSheetHandle` surface.

Test these behaviors:

- `?mode=controlled` selects Controlled on first render;
- an invalid mode renders Uncontrolled and calls `history.replaceState` with the canonical URL;
- selecting a mode updates `?mode=`, changes the keyed preset boundary, and clears old events;
- browser `popstate` restores the matching mode;
- Reset recreates the current fixture without a page reload;
- controlled callbacks update the displayed JSON;
- events retain the newest 50 entries;
- locales switch among the four public dictionaries per instance;
- custom chrome receives typed public slot props/actions;
- `onError` produces an error event;
- an unexpected render failure shows Reset and Reload actions.

- [ ] **Step 2: Run the component test and confirm the page is missing**

```bash
npx vitest run --project component tests/component/docs-playground.test.tsx
```

Expected: import failure for the playground component.

- [ ] **Step 3: Implement the preset boundary and state flow**

In `playground.tsx`, import only:

```ts
import { TegoSheet, type TegoSheetHandle, type WorkbookData } from 'tego-sheet';
import en from 'tego-sheet/locales/en';
import zhCN from 'tego-sheet/locales/zh-cn';
import de from 'tego-sheet/locales/de';
import nl from 'tego-sheet/locales/nl';
import 'tego-sheet/styles.css';
```

Use a keyed boundary such as `key={`${mode}:${resetRevision}`}`. Implement Controlled and Uncontrolled as separate child components so a mounted `TegoSheet` never changes ownership mode. Read current JSON from public callback values or `TegoSheetHandle.getValue()` only.

URL synchronization rules:

- initialize from `new URLSearchParams(window.location.search)`;
- use Docusaurus history-compatible `pushState` for user mode selection;
- subscribe to `popstate` and clean up the listener;
- use `replaceState` only for invalid/missing canonicalization;
- preserve the `/tego-sheet/playground` pathname and unrelated query parameters.

- [ ] **Step 4: Implement inspector, reset, error boundary, and accessibility**

The inspector must expose mode description, public APIs, guide link, Reset mode, newest-first event list, and read-only formatted JSON with a Copy button. Use native buttons, visible focus states, an `aria-live="polite"` status region for mode/reset/copy outcomes, and labeled controls.

The error boundary provides:

- Reset: increments the boundary key and returns to the default preset if recovery fails;
- Reload: invokes `window.location.reload()`;
- no stack trace in the public UI.

- [ ] **Step 5: Create the Docusaurus client-only page boundary**

`website/src/pages/playground.tsx` must use `@docusaurus/BrowserOnly`:

```tsx
<BrowserOnly fallback={<PlaygroundLoadingState />}>
  {() => <Playground />}
</BrowserOnly>
```

Wrap it in `Layout`, set a descriptive title, and render a stable full-width loading skeleton during SSR.

- [ ] **Step 6: Make component tests pass and prove SSR safety**

```bash
npx vitest run --project component tests/component/docs-playground.test.tsx
npm run typecheck:docs
npm run docs:build
```

Expected: all component tests pass and Docusaurus statically renders `/playground` without `window`, Canvas, or DOM access errors.

- [ ] **Step 7: Commit the interactive playground**

```bash
git add website/src/components/playground website/src/pages/playground.tsx website/src/pages/playground.module.css tests/component/docs-playground.test.tsx
git commit -m "feat(playground): expose five shareable integration modes" \
  -m "Constraint: the static site must render safely before browser APIs exist" \
  -m "Rejected: a live code executor | preset controls cover the requested use cases without arbitrary script risk" \
  -m "Confidence: high" \
  -m "Scope-risk: broad" \
  -m "Directive: access workbook state only through public props, callbacks, and TegoSheetHandle" \
  -m "Tested: playground component suite, docs typecheck, and static Docusaurus build"
```

## Task 7: Complete the home page and responsive visual system

**Files:**

- Modify: `website/src/pages/index.tsx`
- Modify: `website/src/pages/index.module.css`
- Create: `website/src/components/homepage-preview.tsx`
- Modify: `website/src/css/custom.css`
- Modify: `tests/architecture/documentation-site-contract.test.ts`

- [ ] **Step 1: Add a failing home-page content contract**

Assert the home page includes the exact user paths and public import snippet:

```ts
expect(home).toContain('/docs/getting-started/quick-start');
expect(home).toContain('/playground');
expect(home).toContain("from 'tego-sheet'");
expect(home).toContain("import 'tego-sheet/styles.css'");
expect(home).toContain('HomepagePreview');
```

- [ ] **Step 2: Run the focused contract and confirm missing content**

```bash
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: the minimal Task 2 home page lacks the complete snippet and preview.

- [ ] **Step 3: Implement the approved home hierarchy**

Use this order:

1. concise product promise and two CTAs;
2. copyable `npm install tego-sheet` command;
3. minimal controlled React example;
4. four capability statements: typed React state, Canvas performance, public slots/locales, compatible workbook JSON;
5. compact live preview;
6. links to usage guides, API, Playground and GitHub.

`homepage-preview.tsx` must import `TegoSheet` from `tego-sheet` and be wrapped in `BrowserOnly`, with a fixed compact fixture and no inspector or controls.

- [ ] **Step 4: Implement the restrained responsive theme**

Use existing demo colors as CSS custom properties: deep navy workspace, gray-blue borders, bright blue actions. Keep body documentation surfaces light and high contrast; map dark mode through `[data-theme='dark']`. Avoid gradients and decorative motion.

At narrow widths:

- CTAs wrap without overflow;
- preview remains usable but is explicitly a responsive viewport, not a mobile-editing promise;
- Playground inspector moves below the sheet;
- controls remain at least 44px high and keyboard reachable.

- [ ] **Step 5: Verify type, lint, formatting, and strict build**

```bash
npm run format
npm run lint
npm run typecheck:docs
npm run docs:build
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: no lint warnings, no formatting changes after the first pass, no broken routes, and the home contract passes.

- [ ] **Step 6: Commit the site experience**

```bash
git add website/src tests/architecture/documentation-site-contract.test.ts
git commit -m "feat(docs): present a focused spreadsheet learning path" \
  -m "Constraint: documentation readability takes priority over decorative marketing UI" \
  -m "Rejected: card-heavy animated landing patterns | they distract from code, API, and the live sheet" \
  -m "Confidence: high" \
  -m "Scope-risk: moderate" \
  -m "Directive: keep full spreadsheet experimentation on the dedicated Playground page" \
  -m "Tested: format, lint, docs typecheck, strict build, and home architecture contract"
```

## Task 8: Add browser and visual regression coverage for the site

**Files:**

- Create: `playwright.docs.config.ts`
- Create: `playwright.docs-visual.config.ts`
- Create: `tests/docs/docs.spec.ts`
- Create: `tests/docs-visual/docs-visual.spec.ts`
- Create: `tests/docs-visual/docs-visual.spec.ts-snapshots/*-darwin.png`
- Modify: `tests/architecture/visual-regression-contract.test.ts`
- Modify: `tests/package/repository-policy.test.mjs`

- [ ] **Step 1: Add failing policy assertions for both Playwright configs**

Assert:

- docs browser base URL is `http://127.0.0.1:4175/tego-sheet/`;
- docs visual base URL is `http://127.0.0.1:4176/tego-sheet/`;
- both web servers run a fresh `docs:build` before `docs:serve`;
- docs visual output uses a distinct report/results directory;
- docs visual config uses the existing macOS screenshot policy and deterministic font/animation settings.

- [ ] **Step 2: Run policy tests and confirm the configs are missing**

```bash
node --test tests/package/repository-policy.test.mjs
npx vitest run --project architecture tests/architecture/visual-regression-contract.test.ts
```

Expected: missing config failures.

- [ ] **Step 3: Create the docs browser configuration and tests**

`playwright.docs.config.ts` uses Chromium desktop plus one touch/narrow project, `reuseExistingServer: !process.env.CI`, and:

```ts
webServer: {
  command:
    'npm run docs:build && npm run docs:serve -- --host 127.0.0.1 --port 4175',
  url: 'http://127.0.0.1:4175/tego-sheet/',
  timeout: 180_000,
  reuseExistingServer: !process.env.CI,
},
```

`tests/docs/docs.spec.ts` must:

- open the project subpath home and follow Docs, API and Playground links;
- verify stylesheet/script requests return no 404s;
- switch through all five presets and assert `?mode=` values;
- edit one real cell in uncontrolled and controlled modes using the same stable Canvas interaction helpers as existing browser tests;
- assert event and JSON inspector changes;
- verify Reset, reload, forward and back behavior;
- test a narrow viewport with the inspector below the sheet.

- [ ] **Step 4: Create the docs visual configuration and specs**

Follow `playwright.visual.config.ts` for macOS-only baseline naming, fixed color scheme, motion reduction, font wait, thresholds, retries and artifact paths. Capture only:

- home desktop;
- Quick Start desktop;
- Playground desktop in Controlled mode;
- Playground narrow in Uncontrolled mode.

Mask timestamps/sequence values in the event inspector, seed fixed data, and wait for `document.fonts.ready` plus the sheet-ready signal before screenshots.

- [ ] **Step 5: Run browser tests, generate baselines, and apply the visual gate**

```bash
npm run test:docs
npm run test:docs-visual -- --update-snapshots
npm run test:docs-visual
```

Inspect all four baselines at original resolution. Invoke `visual-verdict` for the home, docs, desktop Playground and narrow Playground captures before accepting them. Fix clipping, unreadable contrast, layout overflow, unstable rendering, or accidental animation; regenerate only affected baselines and rerun the verdict.

- [ ] **Step 6: Run regression-policy tests**

```bash
node --test tests/package/repository-policy.test.mjs
npx vitest run --project architecture tests/architecture/visual-regression-contract.test.ts
```

Expected: config/policy tests pass and all baselines are tracked under the macOS naming convention.

- [ ] **Step 7: Commit browser coverage and reviewed baselines**

```bash
git add playwright.docs.config.ts playwright.docs-visual.config.ts tests/docs tests/docs-visual tests/architecture/visual-regression-contract.test.ts tests/package/repository-policy.test.mjs
git commit -m "test(docs): lock project-path behavior and responsive visuals" \
  -m "Constraint: Pages correctness depends on the /tego-sheet/ subpath and deterministic macOS rendering" \
  -m "Rejected: snapshotting every API page | strict generation and link checks cover structure without brittle image volume" \
  -m "Confidence: high" \
  -m "Scope-risk: moderate" \
  -m "Directive: update baselines only after a recorded visual-verdict review" \
  -m "Tested: docs browser suite, docs visual suite, and regression-policy tests"
```

## Task 9: Gate and deploy the tested site through GitHub Actions

**Files:**

- Modify: `.github/workflows/ci.yml`
- Modify: `tests/package/repository-policy.test.mjs`

- [ ] **Step 1: Add failing exact CI-policy expectations**

Extend the expected job list with `docs-site` and `deploy-pages`. Pin these actions exactly:

```js
'actions/configure-pages@983d7736d9b0ae728b81ab479565c72886d7745b', // v5
'actions/upload-pages-artifact@7b1f4a764d45c48632c6b24a0339c27f5614fb0b', // v4
'actions/deploy-pages@d6db90164ac5ed86f2b6aed7e0febac5b3c0c03e', // v4
```

Assert:

- PRs run `docs-site` but neither upload nor deploy Pages;
- `visual` runs both `npm run test:visual` and `npm run test:docs-visual`;
- `docs-site` runs `npm run typecheck:docs` and `npm run test:docs`;
- `deploy-pages.needs` exactly includes `commit-policy`, `quality`, `vitest`, `package-contract`, `browser`, `visual`, `parity-release`, and `docs-site`;
- only `deploy-pages` has `pages: write` and `id-token: write`;
- deploy runs only for a push to `main` and uses environment `github-pages`.

- [ ] **Step 2: Run repository policy and confirm missing jobs/actions**

```bash
node --test tests/package/repository-policy.test.mjs
```

Expected: deterministic failures for `docs-site`, `deploy-pages`, the three action pins, and docs visual command.

- [ ] **Step 3: Extend the visual job**

After the existing library visual test, run:

```yaml
- name: Run documentation visual regression tests
  run: npm run test:docs-visual
```

Upload its Playwright report on failure with the existing pinned `actions/upload-artifact` action and a distinct artifact name.

- [ ] **Step 4: Add the docs-site job**

Use the existing pinned checkout/setup-node conventions, Node 20, `npm ci`, and Chromium install. Then:

```yaml
- name: Configure Pages
  if: github.event_name == 'push' && github.ref == 'refs/heads/main'
  uses: actions/configure-pages@983d7736d9b0ae728b81ab479565c72886d7745b
- name: Typecheck documentation site
  run: npm run typecheck:docs
- name: Run documentation browser tests
  run: npm run test:docs
- name: Upload Pages artifact
  if: github.event_name == 'push' && github.ref == 'refs/heads/main'
  uses: actions/upload-pages-artifact@7b1f4a764d45c48632c6b24a0339c27f5614fb0b
  with:
    path: website/build
```

Upload the docs Playwright report on failure with the existing pinned general artifact action. The browser test's fresh build is the exact directory uploaded; do not run another untested production build afterward.

- [ ] **Step 5: Add the gated deploy-pages job**

```yaml
deploy-pages:
  if: github.event_name == 'push' && github.ref == 'refs/heads/main'
  needs:
    - commit-policy
    - quality
    - vitest
    - package-contract
    - browser
    - visual
    - parity-release
    - docs-site
  runs-on: ubuntu-latest
  permissions:
    pages: write
    id-token: write
  environment:
    name: github-pages
    url: ${{ steps.deployment.outputs.page_url }}
  steps:
    - name: Deploy GitHub Pages
      id: deployment
      uses: actions/deploy-pages@d6db90164ac5ed86f2b6aed7e0febac5b3c0c03e
```

Keep workflow top-level permissions at `contents: read`.

- [ ] **Step 6: Verify the workflow contract and YAML syntax**

```bash
node --test tests/package/repository-policy.test.mjs
npm run lint
npm run format:check
```

Expected: exact job/action/permission assertions pass, and formatting/lint remain clean.

- [ ] **Step 7: Commit the Pages pipeline**

```bash
git add .github/workflows/ci.yml tests/package/repository-policy.test.mjs
git commit -m "ci: publish only fully verified documentation artifacts" \
  -m "Constraint: deployment must wait for library, parity, docs, browser, and visual gates" \
  -m "Rejected: a separate ungated Pages workflow | it could publish code that failed the release contract" \
  -m "Confidence: high" \
  -m "Scope-risk: broad" \
  -m "Directive: keep Pages write and OIDC permissions isolated to deploy-pages" \
  -m "Tested: exact repository workflow policy, lint, and format checks"
```

## Task 10: Link the public site, run the complete release verification, and publish

**Files:**

- Modify: `readme.md`
- Modify if required by verification: only files introduced in Tasks 1–9

- [ ] **Step 1: Add public documentation links without moving attribution**

In `readme.md`, add a `Documentation` section before the existing final `Ownership and upstream attribution` section. Include:

- Documentation: `https://sealday.github.io/tego-sheet/docs/getting-started/installation`
- API Reference: `https://sealday.github.io/tego-sheet/docs/api`
- Playground: `https://sealday.github.io/tego-sheet/playground`

Keep the existing demo screenshot and keep `Ownership and upstream attribution` as the final section exactly as requested earlier.

- [ ] **Step 2: Run the complete local quality sequence from a clean generated state**

Remove ignored generated directories only, then rebuild from source:

```bash
rm -rf website/build website/.docusaurus website/docs/api
npm ci
npm run format:check
npm run lint
npm run typecheck
npm run typecheck:docs
npm run test
npm run test:ssr
npm run test:package
npm run build
npm run build:demo
npm run docs:build
npm run test:browser
npm run test:docs
npm run test:visual
npm run test:docs-visual
npm run test:parity-release
```

Expected: every command exits 0. If a command fails, fix the smallest responsible implementation slice, rerun the focused failing command, then restart this complete sequence at `npm run format:check`.

- [ ] **Step 3: Audit the produced site and package boundary**

```bash
git status --short
git ls-files website/docs/api website/build website/.docusaurus
rg -n "from ['\"](?:\.\./)*src/|from ['\"].*core/|controller|canvas-engine" website/src website/docs
rg -n "https://sealday.github.io/tego-sheet|/tego-sheet/" website/build/index.html website/build/playground.html
```

Expected:

- no generated file is tracked;
- no website source imports private paths;
- built pages contain the correct project-path assets/links;
- `git status` contains only the intended README change before the final commit.

- [ ] **Step 4: Commit the public entry links**

```bash
git add readme.md
git commit -m "docs: make the learning site discoverable" \
  -m "Constraint: ownership and upstream attribution remains the README's final section" \
  -m "Rejected: replacing the local demo instructions | local and hosted verification serve different needs" \
  -m "Confidence: high" \
  -m "Scope-risk: narrow" \
  -m "Tested: complete local quality, browser, visual, package, docs, and parity sequence"
```

- [ ] **Step 5: Verify history, cleanliness, and commit policy**

```bash
git status --short
git log --format='%H%n%B%n---' dd0c984..HEAD
npx commitlint --from dd0c984 --to HEAD
```

Expected: clean worktree; the implementation is split across the planned commits; every commit passes Conventional Commits and retains Lore trailers.

- [ ] **Step 6: Push and enable Pages workflow mode idempotently**

```bash
git push origin main
if current_build_type="$(gh api repos/sealday/tego-sheet/pages --jq .build_type 2>/dev/null)"; then
  if [ "$current_build_type" != "workflow" ]; then
    gh api --method PUT repos/sealday/tego-sheet/pages -f build_type=workflow
  fi
else
  gh api --method POST repos/sealday/tego-sheet/pages -f build_type=workflow
fi
```

Do not create a `gh-pages` branch.

- [ ] **Step 7: Monitor the pushed CI to completion**

```bash
run_id="$(gh run list --workflow ci.yml --branch main --limit 1 --json databaseId --jq '.[0].databaseId')"
gh run watch "$run_id" --exit-status
gh run view "$run_id" --json jobs,url,conclusion
```

Expected: all quality jobs and `deploy-pages` conclude successfully. If any job fails, inspect with `gh run view "$run_id" --log-failed`, fix locally, rerun the relevant local check, commit with the same Lore protocol, push, and monitor the replacement run.

- [ ] **Step 8: Smoke-test the deployed Pages site**

Run:

```bash
curl --fail --location --silent --show-error https://sealday.github.io/tego-sheet/ >/dev/null
curl --fail --location --silent --show-error https://sealday.github.io/tego-sheet/docs/api >/dev/null
curl --fail --location --silent --show-error 'https://sealday.github.io/tego-sheet/playground?mode=controlled' >/dev/null
```

Open the deployed Playground in a browser and perform one controlled edit plus Reset. Confirm scripts/styles have no 404s, the event/JSON inspector updates, and the URL remains shareable.

- [ ] **Step 9: Record final evidence**

Report:

- public Pages URL and successful workflow URL;
- commit range created by the implementation;
- exact local verification commands that passed;
- that generated API/build output remains ignored;
- that only public package imports are used by the site;
- any genuine residual risk, limited to GitHub Pages propagation delay if still observable.
