# Task 7 Report: Documentation browser and visual coverage

## Files

- `src/template/render.ts`
  - Replaced iterable Set spread with `Array.from(new Set(...))` so Docusaurus' Babel pass preserves
    empty generated-sheet mappings and the built browser demo resolves its print target.
- `tests/docs/docs.spec.ts`
  - Added canonical deep-link and printing-guide navigation coverage.
  - Added deterministic two-page invoice, keyboard workspace switching, stale output disablement,
    explicit regeneration, invalid JSON, blocked generation with preserved preview, and injected
    output-action click coverage.
  - The action test intercepts all four controls before React handlers run; it asserts labels and
    clickability without opening native Print or creating downloads.
- `tests/docs-visual/docs-visual.spec.ts`
  - Added ready-state cases at `1440 × 1100`, `1024 × 1100`, and `390 × 1000`.
  - Added desktop stale and compile-blocked diagnostic cases with deterministic inputs and scroll.
- `tests/docs-visual/docs-visual.spec.ts-snapshots/`
  - Added five Output Studio baselines.
  - Refreshed the existing desktop Controlled and narrow Uncontrolled Playground baselines to
    include the Task 6 workspace shell.
- `website/docs/guides/printing.md`
  - Linked Output Studio below the introduction.
  - Clarified that preview/browser print/PDF/PNG consume the immutable display list while XLSX
    writes the semantic workbook in the same `GeneratedDocument`.

## RED evidence

1. `npm run docs:build && npm run test:docs -- --grep "Output Studio"`
   - Initial run exposed a stale server from the main checkout on port 4175; it was stopped and the
     focused RED was rerun against this worktree.
2. `npm run test:docs -- --grep "Output Studio opens directly"`
   - **RED:** 2 failed (desktop and narrow), both because the printing guide had no Output Studio
     link.
3. `npm run test:docs -- --grep "Output Studio"`
   - After adding the link, browser integration remained **RED** because the built site produced
     `Print 0 pages` and no preview articles.
   - Diagnostics showed matching `invoice-sheet` source/target IDs but zero resolved targets.
     Docusaurus transpiled `[...new Set(...)]` to `[].concat(new Set(...))`, turning an empty mapping
     into a one-item array containing the Set. The minimal `Array.from(new Set(...))` fix restored
     Node/browser parity.
4. `npm run test:docs-visual -- --grep "Output Studio"`
   - **RED:** 5 failed solely because the approved Output Studio baselines did not exist.

## GREEN evidence

- `npx vitest run --project unit tests/unit/website/output-studio-pipeline.test.ts`
  - **PASS:** 4 tests, including the deterministic two-page pipeline result.
- `npm run docs:build`
  - **PASS:** library and optimized Docusaurus production builds completed.
- `npm run test:docs`
  - **PASS:** 22 tests across desktop and narrow-touch Chromium projects.
- `npm run test:docs-visual -- --update-snapshots`
  - **PASS:** 10 tests; five Output Studio baselines added and two inherited Playground baselines
    refreshed.
- `npm run test:docs-visual`
  - **PASS:** 10 tests without snapshot updates.
- `npm run typecheck`
  - **PASS.**
- `npm run typecheck:docs`
  - **PASS**, including a fresh library build.
- `npm run lint`
  - **PASS:** zero warnings.
- `npm run format:check`
  - **PASS:** 648 files checked.
- `git diff --check`
  - **PASS.**

## Commits

- Implementation and baselines: `caeb71aa0f0f16fda77a189e204aa74f08720f5f`

## Risks / notes

- Docs browser and visual configurations run Chromium only; broader browser support remains covered
  by the repository's separate browser suites.
- Native Print is intentionally never invoked by these tests. The injected capture test verifies
  the four output controls without OS dialogs or download side effects.
- Output encoder internals are not duplicated here; component and adapter tests remain responsible
  for PDF/PNG/XLSX payload details.
- The one-line renderer scope expansion was required because the production documentation bundle,
  unlike direct Node execution, lost all print targets after Babel transformed the Set spread.
