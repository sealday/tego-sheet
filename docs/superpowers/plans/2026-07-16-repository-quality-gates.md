# Repository Quality Gates Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace ESLint with Oxlint and Oxfmt, enforce Conventional Commits locally and in GitHub Actions, add complete CI coverage, and finish the README with a real demo screenshot and final ownership attribution.

**Architecture:** Repository policy is locked by a dependency-free Node test before configuration changes. Local enforcement uses pinned Oxc tools plus Husky/Commitlint, while one GitHub Actions workflow separates fast pull-request gates from the indivisible main/manual parity release. Documentation changes are isolated from runtime source so the React API and workbook behavior remain unchanged.

**Tech Stack:** npm, Node.js 20/22/24, Oxlint 1.74.0, Oxfmt 0.59.0, Husky 9.1.7, Commitlint 21.2.x, GitHub Actions, Vitest, Playwright, Vite 8.

---

## File map

- Create `tests/package/repository-policy.test.mjs`: dependency-free assertions for toolchain, hooks, workflow, Playwright commands, README ordering, and screenshot tracking.
- Modify `scripts/test-package.mjs`: include the repository-policy test in the existing package-contract runner.
- Create `.oxlintrc.json`: native JavaScript, TypeScript, React Hooks, and React Refresh lint policy.
- Create `.oxfmtrc.json`: repository formatter options and immutable-evidence exclusions.
- Create `commitlint.config.js`: Conventional Commits parser policy.
- Create `.husky/pre-commit`: read-only formatting and lint gate.
- Create `.husky/commit-msg`: Commitlint gate for the proposed message.
- Modify `package.json` and `package-lock.json`: remove ESLint, pin the replacement toolchain, add scripts, and declare npm/Node support.
- Delete `eslint.config.js`: remove the obsolete linter configuration.
- Modify `tests/architecture/public-surface.test.ts`: inspect the new configuration files instead of the removed ESLint file.
- Modify `playwright.config.ts` and `playwright.visual.config.ts`: use npm-native Vite startup and add separate HTML report output.
- Create `.github/workflows/ci.yml`: commit policy, quality, Vitest, package matrix, browser, visual, and restricted parity-release jobs.
- Create `docs/assets/tego-sheet-demo.png`: deterministic 1440×900 screenshot of the full-screen Vite workbench.
- Modify `readme.md`: place the screenshot after the introduction and ownership/upstream attribution last.

### Task 1: Lock the repository policy with a failing test

**Files:**
- Create: `tests/package/repository-policy.test.mjs`
- Modify: `scripts/test-package.mjs`

- [ ] **Step 1: Create the dependency-free repository-policy test**

Write tests using `node:test`, `node:assert/strict`, `existsSync`, `readFileSync`, and `execFileSync`. The assertions must require:

```js
assert.equal(existsSync(new URL('eslint.config.js', root)), false);
assert.equal(pkg.devDependencies.oxlint, '1.74.0');
assert.equal(pkg.devDependencies.oxfmt, '0.59.0');
assert.equal(pkg.devDependencies.husky, '9.1.7');
assert.equal(pkg.devDependencies['@commitlint/cli'], '21.2.1');
assert.equal(pkg.devDependencies['@commitlint/config-conventional'], '21.2.0');
assert.equal(pkg.packageManager, 'npm@11.13.0');
assert.equal(pkg.engines.node, '>=20.19.0');
assert.equal(pkg.scripts.lint, 'oxlint --deny-warnings .');
assert.equal(pkg.scripts['lint:fix'], 'oxlint --fix .');
assert.equal(pkg.scripts.format, 'oxfmt --write .');
assert.equal(pkg.scripts['format:check'], 'oxfmt --check .');
assert.equal(pkg.scripts.prepare, 'husky');
```

Parse `.oxlintrc.json` and `.oxfmtrc.json`, then assert the React rules and immutable-evidence ignore patterns exactly:

```js
assert.deepEqual(oxlint.plugins, ['eslint', 'typescript', 'unicorn', 'oxc', 'react']);
assert.equal(oxlint.rules['react/rules-of-hooks'], 'error');
assert.equal(oxlint.rules['react/exhaustive-deps'], 'error');
assert.deepEqual(oxlint.rules['react/only-export-components'], [
  'error',
  { allowConstantExport: true },
]);
assert.equal(oxfmt.singleQuote, true);
assert.equal(oxfmt.printWidth, 100);
assert.deepEqual(oxfmt.ignorePatterns, [
  'docs/superpowers/**',
  'tests/parity/fixtures/**',
  'tests/parity/legacy/**',
  'tests/visual/__snapshots__/**',
  'tests/visual/fonts/**',
]);
```

Read both Husky hook files and `commitlint.config.js`, asserting the exact commands from the design. Read both Playwright configs and assert they contain `npm exec -- vite`, do not contain `pnpm`, and define `playwright-report/browser` or `playwright-report/visual` HTML reporters. Read `.github/workflows/ci.yml` and assert the seven job IDs, three triggers, `permissions: contents: read`, the Node 20/22/24 matrix, artifact upload actions, and the main/manual condition on `parity-release`.

Finish by checking the README and screenshot:

```js
const screenshot = 'docs/assets/tego-sheet-demo.png';
assert.equal(existsSync(new URL(screenshot, root)), true);
assert.ok(execFileSync('git', ['ls-files', '--error-unmatch', screenshot], {
  cwd: rootPath,
  stdio: 'ignore',
}) === undefined);
assert.match(readme, /!\[Tego Sheet interactive workbench\]\(docs\/assets\/tego-sheet-demo\.png\)/);
assert.equal(readme.trimEnd().lastIndexOf('## Ownership and upstream attribution'), readme.trimEnd().indexOf('## Ownership and upstream attribution'));
assert.ok(readme.trimEnd().endsWith('Third-party assets that carry their own notices remain subject to their respective licenses.'));
```

- [ ] **Step 2: Add the new test to package orchestration**

Insert `'tests/package/repository-policy.test.mjs'` into the `runNodeTests` array in `scripts/test-package.mjs` after `quality-gates.test.mjs`.

- [ ] **Step 3: Run the focused test and record the expected RED result**

Run: `node --test tests/package/repository-policy.test.mjs`

Expected: FAIL because ESLint still exists and the replacement configuration, hooks, workflow, and screenshot do not exist.

- [ ] **Step 4: Commit the failing contract**

```bash
git add tests/package/repository-policy.test.mjs scripts/test-package.mjs
git commit -m "test: lock repository quality policy" -m "Constraint: Preserve runtime behavior while replacing repository infrastructure
Confidence: high
Scope-risk: narrow
Tested: node --test tests/package/repository-policy.test.mjs (expected failure)
Not-tested: replacement toolchain and CI implementation"
```

### Task 2: Replace ESLint and install local repository gates

**Files:**
- Create: `.oxlintrc.json`
- Create: `.oxfmtrc.json`
- Create: `commitlint.config.js`
- Create: `.husky/pre-commit`
- Create: `.husky/commit-msg`
- Modify: `package.json`
- Modify: `package-lock.json`
- Modify: `tests/architecture/public-surface.test.ts`
- Delete: `eslint.config.js`

- [ ] **Step 1: Replace dependencies with exact pinned versions**

Run:

```bash
npm uninstall @eslint/js eslint eslint-plugin-react-hooks eslint-plugin-react-refresh typescript-eslint
npm install --save-dev --save-exact oxlint@1.74.0 oxfmt@0.59.0 husky@9.1.7 @commitlint/cli@21.2.1 @commitlint/config-conventional@21.2.0
```

- [ ] **Step 2: Add package scripts and runtime metadata**

Set these exact `package.json` properties without changing package exports or peer dependencies:

```json
{
  "packageManager": "npm@11.13.0",
  "engines": {
    "node": ">=20.19.0"
  },
  "scripts": {
    "lint": "oxlint --deny-warnings .",
    "lint:fix": "oxlint --fix .",
    "format": "oxfmt --write .",
    "format:check": "oxfmt --check .",
    "prepare": "husky"
  }
}
```

This repository currently runs npm 11.13.0, so the package manager declaration is pinned to that exact version.

- [ ] **Step 3: Add the Oxlint configuration**

Create `.oxlintrc.json`:

```json
{
  "$schema": "./node_modules/oxlint/configuration_schema.json",
  "plugins": ["eslint", "typescript", "unicorn", "oxc", "react"],
  "categories": {
    "correctness": "error"
  },
  "rules": {
    "react/rules-of-hooks": "error",
    "react/exhaustive-deps": "error",
    "react/only-export-components": ["error", { "allowConstantExport": true }]
  },
  "ignorePatterns": [
    ".nyc_output/**",
    ".worktrees/**",
    "coverage/**",
    "demo-dist/**",
    "dist/**",
    "docs/**",
    "node_modules/**",
    "playwright-report/**",
    "test-results/**"
  ]
}
```

- [ ] **Step 4: Add the Oxfmt configuration**

Create `.oxfmtrc.json`:

```json
{
  "$schema": "./node_modules/oxfmt/configuration_schema.json",
  "singleQuote": true,
  "printWidth": 100,
  "ignorePatterns": [
    "docs/superpowers/**",
    "tests/parity/fixtures/**",
    "tests/parity/legacy/**",
    "tests/visual/__snapshots__/**",
    "tests/visual/fonts/**"
  ]
}
```

- [ ] **Step 5: Add Commitlint and direct Husky hooks**

Create `commitlint.config.js`:

```js
export default {
  extends: ['@commitlint/config-conventional'],
};
```

Create executable `.husky/pre-commit`:

```sh
npm run format:check && npm run lint
```

Create executable `.husky/commit-msg`:

```sh
npx --no -- commitlint --edit "$1"
```

Run: `chmod +x .husky/pre-commit .husky/commit-msg`

Run: `npm run prepare` and verify `git config --get core.hooksPath` prints `.husky/_`.

- [ ] **Step 6: Remove ESLint and update the architecture guard**

Delete `eslint.config.js`. In `tests/architecture/public-surface.test.ts`, replace `'eslint.config.js'` in the configuration-file loop with `'.oxlintrc.json'` and `'.oxfmtrc.json'`.

- [ ] **Step 7: Format, lint, and run focused tests**

Run:

```bash
npm run format
npm run format:check
npm run lint
npm run typecheck
node --test tests/package/repository-policy.test.mjs
npm exec -- vitest run --project architecture
```

Expected: formatting, lint, typecheck, and architecture pass; repository-policy remains RED only for the still-missing workflow, npm-native Playwright reporters, and README screenshot/order.

- [ ] **Step 8: Prove the local commit policy**

Run:

```bash
printf 'not conventional\n' | npm exec -- commitlint
printf 'build: replace repository quality tooling\n' | npm exec -- commitlint
```

Expected: first command exits nonzero; second exits zero.

- [ ] **Step 9: Commit the local toolchain**

```bash
git add package.json package-lock.json .oxlintrc.json .oxfmtrc.json commitlint.config.js .husky tests/architecture/public-surface.test.ts eslint.config.js
git commit -m "build: enforce repository quality locally" -m "Constraint: Oxlint owns diagnostics and Oxfmt owns deterministic formatting
Rejected: Keep Prettier beside Oxfmt | one formatter avoids conflicting output
Confidence: high
Scope-risk: moderate
Tested: format check; Oxlint; typecheck; architecture tests; Commitlint probes
Not-tested: GitHub-hosted hooks and workflow jobs"
```

### Task 3: Add GitHub Actions and make Playwright npm-native

**Files:**
- Create: `.github/workflows/ci.yml`
- Modify: `playwright.config.ts`
- Modify: `playwright.visual.config.ts`

- [ ] **Step 1: Change both Playwright web servers and reporters**

Use `command: 'npm exec -- vite --config …'` in both configs. Add an HTML reporter before the parity reporter:

```ts
['html', { open: 'never', outputFolder: 'playwright-report/browser' }],
```

and in the visual config:

```ts
['html', { open: 'never', outputFolder: 'playwright-report/visual' }],
```

- [ ] **Step 2: Create the CI workflow**

Create `.github/workflows/ci.yml` with:

```yaml
name: CI

on:
  pull_request:
    branches: [main]
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read

concurrency:
  group: ci-${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: true

env:
  CI: 1

jobs:
  commit-policy:
    if: github.event_name != 'workflow_dispatch'
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@v6
        with:
          fetch-depth: 0
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - name: Validate introduced commits
        env:
          BASE_SHA: ${{ github.event.pull_request.base.sha || github.event.before }}
          HEAD_SHA: ${{ github.event.pull_request.head.sha || github.sha }}
        run: npm exec -- commitlint --from "$BASE_SHA" --to "$HEAD_SHA" --verbose

  quality:
    runs-on: ubuntu-latest
    timeout-minutes: 15
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm run format:check
      - run: npm run lint
      - run: npm run typecheck
      - run: npm run build
      - run: npm run build:demo

  vitest:
    runs-on: ubuntu-latest
    timeout-minutes: 15
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm test

  package-contract:
    runs-on: ubuntu-latest
    timeout-minutes: 20
    strategy:
      fail-fast: false
      matrix:
        node: [20, 22, 24]
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: ${{ matrix.node }}
          cache: npm
      - run: npm ci
      - run: npm run test:ssr
      - run: npm run test:package

  browser:
    runs-on: ubuntu-latest
    timeout-minutes: 30
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm exec -- playwright install --with-deps
      - run: npm run test:browser
      - if: ${{ !cancelled() }}
        uses: actions/upload-artifact@v5
        with:
          name: browser-report
          path: |
            playwright-report/browser
            test-results/playwright
          if-no-files-found: ignore

  visual:
    runs-on: ubuntu-latest
    timeout-minutes: 30
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm exec -- playwright install --with-deps chromium
      - run: npm run test:visual
      - if: ${{ !cancelled() }}
        uses: actions/upload-artifact@v5
        with:
          name: visual-report
          path: |
            playwright-report/visual
            test-results/playwright-visual
          if-no-files-found: ignore

  parity-release:
    if: github.event_name == 'workflow_dispatch' || (github.event_name == 'push' && github.ref == 'refs/heads/main')
    runs-on: ubuntu-latest
    timeout-minutes: 45
    steps:
      - uses: actions/checkout@v6
      - uses: actions/setup-node@v6
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm exec -- playwright install --with-deps
      - run: npm run test:parity-release
      - if: ${{ !cancelled() }}
        uses: actions/upload-artifact@v5
        with:
          name: parity-release-evidence
          path: |
            test-results/parity
            test-results/playwright
            test-results/playwright-visual
            playwright-report/browser
            playwright-report/visual
          if-no-files-found: ignore
```

- [ ] **Step 3: Run focused policy and configuration tests**

Run:

```bash
npm run format
node --test tests/package/repository-policy.test.mjs
npm exec -- vitest run tests/parity/evidence-configuration.test.ts
npm run lint
```

Expected: the new workflow and Playwright assertions pass; repository-policy remains RED only for the README screenshot/order.

- [ ] **Step 4: Smoke-test both npm-native Playwright servers**

Run:

```bash
npm exec -- playwright test tests/browser/workbook.spec.ts --project chromium-desktop
npm exec -- playwright test --config playwright.visual.config.ts --project desktop-dpr1 --grep "default workbook"
```

Expected: both tests pass and produce their separate HTML report directories.

- [ ] **Step 5: Commit CI and Playwright portability**

```bash
git add .github/workflows/ci.yml playwright.config.ts playwright.visual.config.ts
git commit -m "ci: enforce complete repository verification" -m "Constraint: Pull requests need full behavior checks without duplicating the release provenance run
Rejected: Run parity release on every pull request | browser and visual jobs already provide equivalent PR feedback
Confidence: high
Scope-risk: moderate
Directive: Keep the parity release command indivisible
Tested: policy test; parity configuration test; Chromium browser and visual smoke tests
Not-tested: GitHub-hosted Firefox and WebKit jobs"
```

### Task 4: Add the demo screenshot and finalize README attribution

**Files:**
- Create: `docs/assets/tego-sheet-demo.png`
- Modify: `readme.md`

- [ ] **Step 1: Start a deterministic local demo server**

Run in a persistent terminal:

```bash
npm run dev -- --host 127.0.0.1 --port 4175
```

Expected: Vite serves the full-screen workbench at `http://127.0.0.1:4175/`.

- [ ] **Step 2: Capture the full-screen demo**

Create `docs/assets`, then run:

```bash
npm exec -- playwright screenshot --browser chromium --viewport-size "1440,900" --wait-for-timeout 1500 http://127.0.0.1:4175/ docs/assets/tego-sheet-demo.png
```

Expected: a non-empty 1440×900 PNG showing the workbench controls and spreadsheet.

- [ ] **Step 3: Add the screenshot after the introduction**

Insert immediately after the opening paragraph:

```markdown
![Tego Sheet interactive workbench](docs/assets/tego-sheet-demo.png)
```

- [ ] **Step 4: Move ownership attribution to the end without rewriting it**

Move the complete existing `## Ownership and upstream attribution` section, including all three paragraphs, to the end of `readme.md`. The file must end with:

```markdown
Tego Sheet is a separate project and is not affiliated with or endorsed by the upstream project. Third-party assets that carry their own notices remain subject to their respective licenses.
```

- [ ] **Step 5: Verify the image and documentation contract**

Run:

```bash
file docs/assets/tego-sheet-demo.png
sips -g pixelWidth -g pixelHeight docs/assets/tego-sheet-demo.png
git add docs/assets/tego-sheet-demo.png readme.md
node --test tests/package/repository-policy.test.mjs
npm run format:check
```

Expected: PNG is 1440×900, repository-policy is fully GREEN, and formatting passes.

- [ ] **Step 6: Commit README presentation**

```bash
git commit -m "docs: show the interactive React workbench" -m "Constraint: Preserve existing upstream attribution wording while moving it after usage documentation
Confidence: high
Scope-risk: narrow
Tested: image dimensions; repository policy; format check
Not-tested: GitHub README rendering"
```

### Task 5: Run complete verification, repair only evidenced failures, and publish

**Files:**
- Modify only files implicated by a failing verification command.

- [ ] **Step 1: Run static and build verification**

Run:

```bash
npm run format:check
npm run lint
npm run typecheck
npm run build
npm run build:demo
```

Expected: all commands exit zero with no warnings denied by Oxlint.

- [ ] **Step 2: Run the complete non-browser test suite and package contracts**

Run:

```bash
npm test
npm run test:ssr
npm run test:package
```

Expected: all Vitest projects, SSR tests, packed consumers, metadata checks, and repository policy pass.

- [ ] **Step 3: Run browser, visual, and indivisible parity verification**

Run:

```bash
npm exec -- playwright install chromium firefox webkit
npm run test:browser
npm run test:visual
npm run test:parity-release
```

Expected: all browser projects, visual snapshots, and parity provenance/evidence gates pass without updating baselines.

- [ ] **Step 4: Inspect repository and workflow integrity**

Run:

```bash
git diff --check
git status --short
git log --format='%h %s' c7cc1df..HEAD
gh workflow view ci.yml --yaml
```

Expected: no unstaged changes, every new commit has a Conventional Commit header, and GitHub CLI accepts and displays the workflow YAML.

- [ ] **Step 5: Repair any verified failure in a focused commit**

For each failure, change only the implicated configuration or test, rerun the failing command and its nearest parent suite, then commit with `fix:` plus Lore trailers describing the evidence. Do not change runtime source or visual baselines unless a test proves the infrastructure change exposed a real pre-existing defect.

- [ ] **Step 6: Push the verified main branch**

Run:

```bash
git push origin main
gh run list --workflow ci.yml --branch main --limit 1
```

Expected: push succeeds and a new CI run appears for `main`.

- [ ] **Step 7: Monitor the GitHub Actions run to completion**

Run: `gh run watch --exit-status <run-id>`

Expected: every required job passes. If a GitHub-only failure occurs, inspect it with `gh run view <run-id> --log-failed`, apply the smallest focused repair, rerun local validation, commit, push, and watch the replacement run.
