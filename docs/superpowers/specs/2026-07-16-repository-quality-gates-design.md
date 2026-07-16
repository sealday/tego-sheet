# Repository Quality Gates Design

## Context

Tego Sheet currently uses ESLint, has no repository-owned GitHub Actions workflow, and has no local Git-hook enforcement. The package is npm-based, but both Playwright configurations start Vite through `pnpm`, which is not guaranteed on a clean GitHub runner. The README also lacks a full demo screenshot and places ownership attribution before the usage documentation.

This change replaces the lint and format toolchain, adds local and remote policy enforcement, makes browser tests npm-native, and improves the README without changing the public React API or spreadsheet behavior.

## Goals

1. Replace ESLint and its plugins with Oxlint while preserving JavaScript, TypeScript, React Hooks, and React Refresh correctness coverage.
2. Use Oxfmt as the only formatter; do not install Prettier.
3. Enforce Conventional Commits locally with Commitlint and Husky and remotely in GitHub Actions.
4. Run complete, reproducible CI for source quality, package compatibility, unit/component/architecture tests, browser behavior, visual snapshots, and the existing parity release gate.
5. Add a real full-screen demo screenshot to the README and move `Ownership and upstream attribution` to the final section.

## Non-goals

- Publishing an npm release.
- Automatically rewriting or staging files inside Git hooks.
- Adding `lint-staged` or another staged-file orchestrator.
- Changing runtime source behavior, public exports, workbook data, visual baselines, or parity evidence.
- Adding a coverage percentage threshold before a separate baseline and policy decision.

## Toolchain

The repository will pin the following development tools and record their exact transitive graph in `package-lock.json`:

- `oxlint@1.74.0`
- `oxfmt@0.59.0`
- `husky@9.1.7`
- `@commitlint/cli@20.5.3`
- `@commitlint/config-conventional@20.5.3`

Commitlint 20.5.3 is the newest release line compatible with the declared Node 20 floor; Commitlint 21.x requires Node 22.12.0 or newer.

The ESLint configuration and the `@eslint/js`, `eslint`, `eslint-plugin-react-hooks`, `eslint-plugin-react-refresh`, and `typescript-eslint` dependencies will be removed completely.

`.oxlintrc.json` will be generated from the existing flat ESLint configuration and reviewed so the enabled rules cover the current JavaScript/TypeScript recommendations plus React Hooks and React Refresh constraints. `lint` will run Oxlint with warnings denied, and `lint:fix` will apply only safe Oxlint fixes.

The migration is an explicit preset mapping rather than a bare category switch:

- `@eslint/js` recommended contributes 65 rules. Oxlint's native `correctness` category covers 53; the JavaScript override adds the 10 supported rules that the category omits. Oxlint 1.74.0 has no named `no-dupe-args` or `no-octal` rule. The explicit `no-redeclare` rule still rejects duplicate parameters, while legacy octal literals are the one unsupported named-rule gap.
- `typescript-eslint` recommended contributes 24 effective core and TypeScript rules. Twelve are native `correctness` coverage and twelve are explicit in the TypeScript override. Oxlint does not expose TypeScript-prefixed `no-array-constructor`, `no-unused-expressions`, or `no-unused-vars`; their core Oxlint equivalents provide the same repository coverage.
- `eslint-plugin-react-hooks` recommended contributes `rules-of-hooks`, `exhaustive-deps`, and fourteen React Compiler diagnostics. Oxlint exposes the first two individually and the compiler diagnostics through the native `react/react-compiler` aggregate, which is enabled for React source and tests.
- The Vite React Refresh preset maps to `react/only-export-components` with `allowConstantExport: true`, scoped only to `src/**/*.tsx`.

The compiler aggregate can report a broader diagnostic at a single source line than the granular ESLint rules. Any necessary exception must therefore remain line-local and explain the timing invariant it preserves; the aggregate is not disabled by file or configuration override.

`.oxfmtrc.json` will use single quotes for JavaScript/TypeScript, a 100-column print width, and repository-wide formatting. It will exclude immutable or historical evidence under `docs/superpowers`, `tests/parity/fixtures`, `tests/parity/legacy`, `tests/visual/__snapshots__`, and `tests/visual/fonts`. `format` writes changes; `format:check` is read-only and is the CI and hook gate. Oxfmt, not Oxlint, owns formatting.

The package will declare npm as its package manager and Node `>=20.19.0`, matching Vite 8's supported runtime floor and the CI compatibility matrix.

## Local Git policy

`commitlint.config.js` will extend `@commitlint/config-conventional` without weakening its parser or type rules.

Husky v9 will be installed through the `prepare` script and will own two direct POSIX hook files:

- `.husky/pre-commit`: `npm run format:check && npm run lint`
- `.husky/commit-msg`: `npx --no -- commitlint --edit "$1"`

The pre-commit hook is intentionally read-only. Developers use `npm run format` and `npm run lint:fix` explicitly, then review and stage the resulting changes. CI repeats every hook policy because hooks can be bypassed.

## GitHub Actions

A single `.github/workflows/ci.yml` workflow will run on pull requests targeting `main`, pushes to `main`, and manual dispatch. It will use read-only repository permissions, npm lockfile caching, bounded job timeouts, and concurrency cancellation by workflow and ref.

The workflow contains these jobs:

1. `commit-policy`
   - Uses a full checkout.
   - On pull requests, validates every commit from the base SHA through the head SHA.
   - On pushes, validates only the new push range so the pre-policy history is not retroactively rejected.
   - Skips manual dispatch because it introduces no commit range.
2. `quality`
   - Runs `npm ci`, Oxfmt check, Oxlint, typecheck, library build, and demo build on Node 24.
3. `vitest`
   - Runs the complete Vitest configuration rather than the narrower `test:unit` alias.
4. `package-contract`
   - Uses Node 20, 22, and 24.
   - Runs SSR and packed-package consumer tests to validate types, ESM, CommonJS, exports, and clean Vite consumption.
5. `browser`
   - Installs all Playwright browsers and system dependencies.
   - Runs the browser suite with one worker and uploads reports and traces when the job is not cancelled.
6. `visual`
   - Installs Chromium and its system dependencies.
   - Runs the visual snapshot suite and uploads reports, diffs, and traces when the job is not cancelled.
7. `parity-release`
   - Runs only for `main` pushes and manual dispatch.
   - Installs all Playwright browsers and executes the existing indivisible parity-release command so Vitest, browser, visual, and manifest evidence share one provenance context.
   - Uploads parity evidence and Playwright output when the job is not cancelled.

The Playwright web servers will use `npm exec -- vite`, removing the undeclared pnpm dependency. Browser and visual configs will add non-opening HTML reporters with separate output folders for CI artifacts.

## README and screenshot

A new tracked image at `docs/assets/tego-sheet-demo.png` will be captured from the repository's full-screen Vite demo at a deterministic desktop viewport after the implementation is complete. The README will display the image immediately after its introductory paragraph.

The existing ownership wording remains intact but moves to the final `## Ownership and upstream attribution` section. No claim about clean-room implementation or absence of adapted upstream logic will be introduced.

## Test-first implementation

Before changing production configuration, repository-policy tests will be added and observed failing. They will assert:

- ESLint configuration and dependencies are absent.
- Oxlint, Oxfmt, Husky, and Commitlint scripts and dependencies are present.
- Commitlint extends the conventional configuration.
- Husky owns the required read-only hooks.
- Playwright no longer invokes pnpm.
- README references a tracked screenshot and ends with the ownership section.
- GitHub Actions contains all required jobs, triggers, minimum permissions, cache setup, Node matrix, report uploads, and the restricted parity-release condition.
- Package manager and Node engine metadata agree with CI.

After the failing policy test is recorded, implementation proceeds in small commits. Verification includes the focused policy test, Oxfmt check, Oxlint, typecheck, full Vitest, builds, SSR/package tests, browser tests, visual tests, parity release, hook behavior, workflow syntax inspection, and a clean Git status.

## Commit and delivery policy

All new commits use Conventional Commit headers and retain the repository's Lore trailers where they add decision context. The expected sequence is:

1. `docs:` approved design.
2. `test:` failing repository-policy contract.
3. `build:` Oxlint, Oxfmt, Commitlint, Husky, and npm metadata.
4. `ci:` GitHub Actions and npm-native Playwright servers.
5. `docs:` demo screenshot and README ordering.
6. Any focused `fix:` commit required by verification.

The completed branch is pushed directly to the existing `main` branch only after local verification succeeds.

## Risks and mitigations

- **Oxlint rule mismatch:** migrate from the current configuration, inspect effective rules, and retain explicit policy tests for React directives.
- **Large formatting diff:** exclude historical and immutable evidence, inspect the formatter diff before committing, and keep formatting separate from behavior changes.
- **Visual CI drift:** use the existing deterministic fonts, viewports, locales, time zone, DPR matrix, and one-worker Playwright configuration.
- **Commitlint rejecting old history:** validate only event-introduced commit ranges, never all repository commits.
- **Expensive parity runs:** keep browser and visual checks required on pull requests, but reserve the duplicate full provenance release for `main` and manual dispatch.
