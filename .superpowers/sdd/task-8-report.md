# Task 8 Report: Integrated verification and cleanup

## Outcome

- All targeted feature, static, package, documentation, full-suite, SSR, and browser visual gates
  completed successfully after two verification-contract fixes.
- The initial automated failure was an architecture assertion that still expected the five
  pre-Output-Studio documentation screenshots. The visual spec now intentionally contains eight
  screenshot calls.
- Final verification also identified Rolldown's slow-plugin advisory in the warning-free static
  gate. The supported per-check switch now suppresses only that non-diagnostic advisory.
- No dependency, production API, or feature implementation changed.

## Verification-only fix

- `tests/architecture/visual-regression-contract.test.ts`
  - Updated the deterministic documentation screenshot count from `5` to `8`.
  - Root cause: commit `caeb71a` added three ready-state Output Studio screenshot calls
    (desktop, intermediate, and narrow) while the architecture count remained at its prior value.
- `vite.config.ts` and `tests/architecture/toolchain-boundaries.test.ts`
  - Set Rolldown `checks.pluginTimings` to `false` and locked the setting with an architecture test.
  - This suppresses only Rolldown's non-diagnostic slow-plugin performance advisory. Build
    warnings and errors retain their normal log handling.

## Exact automated commands and results

### Targeted feature tests

```text
npx vitest run --project unit \
  tests/unit/website/playground-workspace.test.ts \
  tests/unit/website/output-studio-model.test.ts \
  tests/unit/website/output-studio-pipeline.test.ts
```

- PASS: 3 files, 16 tests.

```text
npx vitest run --project component \
  tests/component/docs-playground.test.tsx \
  tests/component/docs-output-studio.test.tsx \
  tests/component/template-designer.test.tsx
```

- PASS: 3 files, 56 tests.
- This coverage includes the same-document output contract, stale result rejection, output abort
  and cleanup on draft changes/regeneration/unmount, accessible busy/error announcements, URL
  workspace behavior, and Spreadsheet regression coverage.

### Static verification

```text
npm run format:check
```

- PASS: 648 files matched repository formatting.

```text
npm run lint
```

- PASS: Oxlint exited 0 with `--deny-warnings` and no findings.

```text
npm run typecheck
```

- PASS: root TypeScript project exited 0 with no diagnostics.

```text
npm run typecheck:docs
```

- PASS: production library/declaration build and website TypeScript project exited 0 with no
  TypeScript diagnostics.
- PASS: the exact command emitted no warnings after narrowly disabling Rolldown's
  `pluginTimings` performance advisory.

### Package and documentation gates

```text
npm run build
```

- PASS: ESM and CommonJS production bundles and declaration files built successfully.

```text
npm run test:package
```

- PASS: 47 of 47 package tests, including clean Vite/NodeNext/CommonJS consumer probes.
- The fixture build printed its existing chunk-size advisory; the package gate exited 0.

```text
npm run docs:build
```

- PASS: Docusaurus client/server compilation and production static generation completed.
- Node printed its existing experimental localStorage advisory during static generation.

```text
npm run test:docs
```

- PASS: 22 of 22 Playwright tests across desktop and narrow-touch projects.
- Verified direct Output Studio linking, keyboard workspace selection, stale/blocked preservation,
  safe stubbed Print behavior, public preset history/reload behavior, Canvas edits, and narrow
  Spreadsheet layout.

```text
npm run test:docs-visual
```

- PASS: 10 of 10 Playwright visual tests.
- Output Studio ready screenshots passed at 1440px, 1024px, and 390px; desktop stale and blocked
  diagnostic screenshots also passed.

### Full suite and SSR

Initial command:

```text
npm test
```

- RED: 1 architecture failure, 1,825 passed, 1 skipped.
- Expected five `toHaveScreenshot` calls but found eight in
  `tests/docs-visual/docs-visual.spec.ts`.

Focused RED-to-GREEN command after the one-line contract correction:

```text
npx vitest run --project architecture \
  tests/architecture/visual-regression-contract.test.ts
```

- PASS: 1 file, 6 tests.

Fresh downstream rerun:

```text
npm test
```

- PASS: 177 files, 1,826 tests; 1 release-evidence test skipped by its environment gate.
- The suite retains existing intentional mount-option console diagnostics and React test `act`
  diagnostics in unrelated legacy component cases; there were no test failures.

```text
npm run test:ssr
```

- PASS: every ESM/CommonJS public entry imported without browser globals.
- PASS: SSR controller epoch project, 1 file and 1 test.

### Commit-hook verification

The verification-fix commit reran:

```text
npm run format:check
npm run lint
```

- PASS: formatting across 648 files and Oxlint with no warnings.

### Final-verifier follow-up

```text
npx vitest run --project architecture \
  tests/architecture/toolchain-boundaries.test.ts
```

- RED before configuration: 1 failed, 3 passed; `checks` was `undefined`.
- PASS after configuration: 1 file, 4 tests.

```text
npm run format:check
npm run lint
npm run typecheck
npm run typecheck:docs
```

- PASS: all four exact static gates exited 0; no static gate emitted a warning.

```text
npm run docs:build
npm run test:docs
npm run test:docs-visual
```

- PASS: production documentation build, 22 of 22 browser tests, and 10 of 10 visual tests.

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

- PASS: 1 file, 24 tests covering the durable workbench, output, cleanup, and focus contracts
  referenced below.

## Browser visual and behavior QA

Served the production Docusaurus output:

```text
npm run docs:serve-static -- --host 127.0.0.1 --port 4188
```

Then drove a real Chromium session against:

```text
http://127.0.0.1:4188/tego-sheet/playground?workspace=output
```

The session used light color scheme and reduced motion. Output action clicks were intercepted at
capture phase, and `window.print` was instrumented, so browser automation never invoked a native
print dialog.

### Responsive inspection

| Viewport               | Document client/scroll width | Print pages | Result                                                     |
| ---------------------- | ---------------------------: | ----------: | ---------------------------------------------------------- |
| Desktop 1440×1100      |                  1440 / 1440 |           2 | Three-column layout; no clipped viewport elements          |
| Intermediate 1024×1100 |                  1024 / 1024 |           2 | Preview-first stacked layout; no clipped viewport elements |
| Narrow 390×1000        |                    390 / 390 |           2 | Single-column layout; no horizontal page overflow          |

The manual QA session produced ignored, transient inspection captures under
`.superpowers/sdd/task-8-evidence/`. They were inspected during Task 8 but are not claimed as
durable repository artifacts.

Durable tracked visual evidence:

- `tests/docs-visual/docs-visual.spec.ts-snapshots/output-studio-ready-desktop-darwin.png`
- `tests/docs-visual/docs-visual.spec.ts-snapshots/output-studio-ready-intermediate-darwin.png`
- `tests/docs-visual/docs-visual.spec.ts-snapshots/output-studio-ready-narrow-darwin.png`
- `tests/docs-visual/docs-visual.spec.ts-snapshots/output-studio-stale-desktop-darwin.png`
- `tests/docs-visual/docs-visual.spec.ts-snapshots/output-studio-blocked-desktop-darwin.png`

Durable automated behavioral evidence:

- `tests/component/docs-output-studio.test.tsx` opens the workbench, verifies drafts remain stale
  until explicit regeneration, and asserts the nested designer's `:focus-visible` outline
  contract.
- `tests/docs/docs.spec.ts` verifies the built site opens the workbench and localizes invalid JSON
  without replacing revision 1.
- `tests/docs-visual/docs-visual.spec.ts` owns the tracked ready, stale, and blocked screenshot
  assertions.

### Behavior evidence

- Initial preview: exactly two `article` elements named `Print page …`, revision 1 visible.
- Same-document invariant: the page states that Preview/Print/PDF/PNG share one exact display list
  and XLSX uses the semantic workbook in the same `GeneratedDocument`; all four ready actions
  derive from that displayed revision.
- Workbench: `aria-expanded` changed `false → true → false`; `#template-workbench` appeared and was
  removed on close.
- Stale state: revision 1 and both preview pages remained visible; Print, PDF, PNG, and XLSX were
  all disabled.
- Invalid JSON: localized alert was exactly
  `Data must be valid JSON before regeneration.`; revision 1 and both preview pages remained.
- Recovery and atomic replacement: a mutation observer saw only
  `{ revision 1, pages 2 } → { revision 2, pages 2 }`; revision 1 was absent after completion.
- Blocked expression: `Generation is blocked. Review the diagnostics.` and the diagnostic list
  were visible; revision 2 and both pages remained; all outputs were disabled.
- Focus: keyboard traversal covered every enabled focusable exposed by the ready, open-workbench
  Output Studio panel. Every focused control matched `:focus-visible` and had a nonzero visible
  outline; the Print action showed a 3px focus ring.
- Print safety: capture-intercepted clicks recorded Print, PDF, PNG, and XLSX in order; instrumented
  native print calls remained `0`, and no print iframe was created.
- Accessibility: one labelled workspace tablist; Output Studio selected; inactive Spreadsheet
  tabpanel hidden; polite status live region; output buttons exposed `aria-busy="false"`; both
  preview articles had accessible Print-page names.
- Deep link/history: direct Output Studio URL canonicalized with `mode=uncontrolled`; clicking
  Spreadsheet and Output Studio pushed the matching workspace URLs; browser Back restored
  Spreadsheet and Forward restored Output Studio.
- Spreadsheet regression: legacy `?mode=controlled` canonicalized to
  `workspace=spreadsheet`; Controlled remained checked, the sheet and 966×670 Canvas were visible,
  and document overflow was zero.

## Commit

- Verification fix: `0f335d2506318cc863d48e66db5ba9d3e1b5e5d8`
  (`fix(docs): harden output studio verification`)
- Initial report: `d1ed9da86abed33ba253d9e70feae268d31c21e7`
  (`docs: record task 8 verification evidence`)
- Final-verifier corrections are recorded in the commit containing this report revision.

## Remaining risks and notes

- Browser QA was performed on macOS Chromium. Platform-scoped deterministic screenshots and the
  automated Docusaurus visual gate provide the release comparison for the same platform.
- Package/docs subprocesses still print their existing chunk-size and experimental localStorage
  advisories outside the warning-free static gate. The exact `typecheck:docs` static gate,
  Oxlint, Oxfmt, and both TypeScript projects reported no warnings or source diagnostics.
- No new dependencies or public APIs were introduced.
