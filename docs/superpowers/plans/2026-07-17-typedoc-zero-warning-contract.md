# TypeDoc Zero-Warning Contract Implementation Plan

**Status:** Completed on 2026-07-17.

**Goal:** Make the real Docusaurus/TypeDoc path warning-free and fail closed on every TypeDoc warning or error while keeping runtime behavior and emitted public declarations unchanged.

**Architecture:** Docusaurus loads TypeDoc through a CommonJS bridge that caches one native ESM module instance. The generator rejects any warning or error after bootstrap, conversion, validation, and output. A production TypeDoc plugin creates documentation-only public projections while preserving the TypeScript declarations; it mutates `TegoSheetProps` only after validating one direct project child, project-owned callback heritage, and the exact inherited callback shape.

**Tech Stack:** TypeScript, Vitest, TypeDoc 0.28, typedoc-plugin-markdown, typedoc-docusaurus-theme, Docusaurus 3, and a Node CommonJS bridge.

---

### Task 1: Lock the real TypeDoc integration contract

- [x] Add a cache-stable native TypeDoc bridge at `website/plugins/strict-typedoc-runtime.cjs`.
- [x] Remove the generator's second runtime TypeDoc load and gate every logger phase unconditionally.
- [x] Enable `treatWarningsAsErrors` in the production Docusaurus configuration.
- [x] Verify the real Docusaurus loader and native ESM plugins share one TypeDoc module instance.
- [x] Cover warning/error cleanup, conversion, validation, and output failure paths.

**Outcome:** The real docs build loads one TypeDoc runtime and reports zero TypeDoc warnings and errors. The child-process integration test retains a measured 30-second upper bound so full-suite resource contention does not cause false failures.

### Task 2: Make public documentation structurally importable

- [x] Assert the exact 55-export TypeDoc root and the direct structure of all fourteen display projections.
- [x] Assert public members have summaries and local named references resolve only to public exports.
- [x] Generate and inspect real Markdown for `TegoSheetProps`, callback headings, and sparse collection summaries.
- [x] Keep all compiler-facing aliases, intersections, and `TegoSheetProps` heritage unchanged.
- [x] Add `website/plugins/public-api-projection.cjs` after proving comment-only inline modifiers do not affect interface heritage expressions.
- [x] Fail closed on missing, duplicate, nested, external, or shape-drifted projection inputs before mutation.

**Outcome:** `TegoSheetProps` Markdown exposes all six callbacks directly without the private helper or inherited labels. `CellsData`, `RowsData`, and `ColsData` use the domain-neutral summary `JSON-compatible entry stored at a sparse decimal index.` Synthetic tests lock helper package, qualified-name, project-package, direct-child, duplicate-child, and callback-shape predicates.

### Task 3: Prove repository and declaration equivalence

- [x] Run the focused public API, strict TypeDoc, and documentation site architecture suites.
- [x] Run the production Docusaurus build and count zero TypeDoc warning/error log entries.
- [x] Run root and docs typechecks, the full repository suite, formatting, lint, and `git diff --check`.
- [x] Build base `2163dd5` and HEAD independently and compare normalized comment-free declaration ASTs.
- [x] Remove `website/docs/api`, `website/build`, and `website/.docusaurus`.

**Outcome:** The final full suite passed 806 tests with one conditional skip. Base and HEAD each emitted 117 declaration files, and the normalized comment-free AST diff was empty.
