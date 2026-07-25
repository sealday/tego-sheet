# Task 5 Report: Output adapters and browser download lifecycle

## Changed files

- `website/src/components/playground/output-download.ts`
  - Added `downloadBlob`, including temporary-anchor removal and microtask Blob URL revocation even
    when the browser click throws.
- `website/src/components/playground/output-studio.tsx`
  - Added injectable `OutputStudioAdapters`.
  - Connected print, PDF, PNG page 1, and XLSX actions to the current `GeneratedDocument`.
  - Used the approved public output adapter entry points and exact option contracts.
  - Added deterministic downloads, independent concurrent busy states, default print-adapter
    disposal, and accessible completion/failure announcements.
  - Kept all output actions disabled for dirty, rendering, and blocked revisions.
  - Review fix: default adapters are now created and owned by effect lifetimes, including React
    StrictMode probes and runtime injection changes; injected adapters are never disposed.
  - Review fix: every output owns an abort controller plus request/revision token. Draft changes,
    render replacement, adapter replacement, and unmount abort active work and suppress late
    downloads or state updates.
  - Review fix: PDF, image, XLSX, and print receive their public `signal` option.
- `website/src/components/playground/output-studio-model.ts`
  - Review fix: replaced the global output result with request-scoped status/message state per
    output kind and committed invoice/title metadata per successful generated revision.
- `website/src/components/playground/playground.module.css`
  - Kept per-action outcomes adjacent to their 44px-minimum initiating controls.
- `tests/component/docs-output-studio.test.tsx`
  - Added adapter identity/options, MIME type, filename, URL lifecycle, concurrent busy state,
    `PRINT_BLOCKED`, XLSX rejection, preview preservation, stale/blocked controls, public import, and
    default disposal coverage.
  - Added StrictMode ownership, injection transitions, adapter-swap cancellation, draft/revision/
    unmount abort guards, stale completion suppression, overlapping retained outcomes, and
    regenerated filename coverage.
- `tests/unit/website/output-studio-model.test.ts`
  - Added request-ID ordering and per-output outcome retention coverage.
- `vitest.config.ts`
  - Added test-only resolution aliases for the existing public PDF, image, and XLSX subpath exports.
- `tsconfig.json`
  - Added matching source aliases so the root test/typecheck program resolves those public subpaths.

## TDD evidence

### RED 1: adapter injection and actions

Command:

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Result: exit 1; 1 failed, 5 passed. The focused test failed because
`adapters.print.print` was called 0 times after clicking `Print 2 pages`.

### GREEN 1

After the minimum adapter wiring and the scoped public-subpath resolver aliases, the same command
passed 6 of 6 tests.

### RED 2: accessible output failures

Command:

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Result: exit 1; 2 failed, 9 passed. Both focused failures found no `role="alert"` for
`PRINT_BLOCKED` and XLSX rejection messages.

### GREEN 2

After rendering failed output messages as assertive alerts, the same command passed 11 of 11 tests.

### RED 3: concurrent independent busy states

Command:

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Result: exit 1; 1 failed, 11 passed. Starting PNG while PDF remained pending incorrectly re-enabled
the PDF button.

### GREEN 3

After tracking busy output kinds independently, the same command passed 12 of 12 tests. Final
coverage additions brought the targeted suite to 14 passing tests.

### Review RED 1: ownership, cancellation, outcomes, and committed metadata

Command:

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Result: exit 1; 7 failed, 14 passed. The focused failures proved:

- the StrictMode effect probe disposed the still-captured default print adapter;
- adapter injection rerenders retained or disposed the wrong adapter;
- output options did not include abort signals and late completions were not guarded;
- a later PDF success removed the retained XLSX alert;
- regenerated data still downloaded with the initial invoice identifier.

### Review GREEN 1

After effect-owned adapters, per-output controllers/tokens, per-kind reducer outcomes, and committed
metadata were implemented, the focused component suite passed 21 of 21 tests and the model suite
passed 9 of 9 tests.

### Review RED 2: adapter replacement busy state

Command:

```text
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Result: exit 1; 1 failed, 22 passed. Replacing injected adapters aborted the old request but left
its PDF control disabled.

### Review GREEN 2

After request cancellation was reflected in per-output reducer state, the focused component suite
passed 23 of 23 tests.

## Verification

- `npx vitest run --project component tests/component/docs-output-studio.test.tsx`
  - PASS: 1 file, 23 tests.
- `npx vitest run --project unit tests/unit/website/output-studio-model.test.ts`
  - PASS: 1 file, 9 tests.
- `npm run lint`
  - PASS: exit 0, no warnings.
- `npx vitest run --project component`
  - PASS: 38 files, 266 tests.
  - Existing unrelated React `act` and mount-only-option warnings remain in other component files;
    the Task 5 targeted suite is clean.
- `npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts`
  - PASS: 1 file, 35 tests, including real website public-package import typechecking.
- `npm run typecheck`
  - PASS: exit 0.
- `npm run build`
  - PASS: Vite production library build and declaration generation.
- `npx tsc --noEmit --project website/tsconfig.json`
  - PASS: exit 0.
- Commit hooks:
  - `npm run format:check` PASS across 648 files.
  - `npm run lint` PASS.

## Commit

- Implementation: `b019cf2` (`feat(docs): connect output studio adapters`)
- Review fixes: `7aa2c51` (`fix(docs): harden output studio lifecycle`)

## Risks

- URL revocation is intentionally queued to the next microtask after the browser accepts or rejects
  the anchor click; tests cover both paths.
- Output completion messages are retained independently until the same action starts again. There
  is no explicit dismiss control in this scoped playground task.
- The Vitest and root TypeScript aliases mirror existing package exports and do not add or change a
  public output API.
