# Task 6 Report: Playground workspace shell

## Changed files

- `website/src/components/playground/playground.tsx`
  - Extracted the existing preset, inspector, reset, recovery, and status behavior into
    `SpreadsheetWorkspace`.
  - Added top-level Spreadsheet and Output Studio tabs.
  - Persisted workspace and mode selection through the Task 1 location reader/writer.
  - Added canonical URL replacement, deep-link initialization, push-state selection, and popstate
    restoration without losing the selected Spreadsheet mode or history announcement.
- `website/src/components/playground/playground.module.css`
  - Added visible, focusable 44px workspace controls.
  - Integrated Output Studio into the shared shell.
  - Applied the approved three-column, 72rem preview-first, and 42rem ordered flex layouts.
- `tests/component/docs-playground.test.tsx`
  - Added legacy URL, Output Studio deep-link, workspace selection, canonical rewrite, and
    cross-workspace history coverage.
  - Preserved the existing Spreadsheet mode remount and live-region history behavior.
- `tests/component/docs-output-studio.test.tsx`
  - Updated the prior responsive assertion to the approved Task 6 42rem ordered flex contract.
- `tests/architecture/documentation-site-contract.test.ts`
  - Added workspace-tab target sizing, responsive layout, and recursive shell module coverage.

## TDD evidence

### RED 1: workspace shell and responsive contract

Commands:

```text
npx vitest run --project component tests/component/docs-playground.test.tsx
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Results:

- Component: exit 1; 6 failed, 16 passed. The failures showed missing workspace tabs and Output
  Studio rendering, legacy mode-only URLs not being rewritten with `workspace=spreadsheet`, mode
  selection still using the old URL writer, and popstate not restoring a workspace.
- Architecture: exit 1; 3 failed, 33 passed. The failures showed missing 44px workspace controls,
  missing approved responsive columns/breakpoints, and no recursive imports of the Task 1 workspace
  model or Output Studio.

### GREEN 1

After the minimal shell extraction, Task 1 URL integration, and scoped CSS changes:

- Playground component suite passed 22 of 22 tests.
- Documentation architecture suite passed 36 of 36 tests.

### RED 2: preserved Spreadsheet history announcement

Command:

```text
npx vitest run --project component tests/component/docs-playground.test.tsx -t "restores a remounted mode"
```

Result: exit 1; 1 failed, 21 skipped. The restored mode was correct, but the live-region status was
empty instead of announcing `Locales restored from browser history`.

### GREEN 2

After tying Spreadsheet status and remount identity to the shell history revision, the focused test
passed. A second RED/GREEN check covered the same announcement when popstate returns from Output
Studio and mounts Spreadsheet again.

## Verification

- `npx vitest run --project component tests/component/docs-playground.test.tsx tests/component/docs-output-studio.test.tsx`
  - PASS: 2 files, 46 tests.
- `npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts`
  - PASS: 1 file, 36 tests.
- `npm run typecheck:docs`
  - PASS: production Vite library build, declaration generation, and website TypeScript check.
- `npx oxlint --deny-warnings website/src/components/playground/playground.tsx tests/component/docs-playground.test.tsx tests/component/docs-output-studio.test.tsx tests/architecture/documentation-site-contract.test.ts`
  - PASS: no warnings.
- `npx oxfmt --check website/src/components/playground/playground.tsx website/src/components/playground/playground.module.css tests/component/docs-playground.test.tsx tests/component/docs-output-studio.test.tsx tests/architecture/documentation-site-contract.test.ts`
  - PASS: all scoped files formatted.
- `git diff --check`
  - PASS.
- Commit hooks:
  - `npm run format:check` PASS across 648 files.
  - `npm run lint` PASS.

## Commit

- Implementation: `7145d5a58e66e56ac5d92bbc60d87ae42b090c1a`
  (`feat(docs): add playground workspace shell`)

## Risks

- Output Studio remains implemented exclusively through the existing public package and output
  subpath imports; Task 6 changes no public printing/output API and adds no dependency.
- Workspace and mode canonicalization intentionally follow the Task 1 helper's parameter ordering
  and preservation behavior.
