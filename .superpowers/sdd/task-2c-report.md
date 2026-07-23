# Task 2C Report: Workbook 2 Runtime

## Outcome

`SpreadsheetDocument` schema 2 is now the single public runtime truth for the controller,
React ingress, change callback, imperative handle, demo, playground, package fixtures, SSR,
and documentation. The legacy sparse workbook model remains private to one operation
projection boundary; `migrateLegacyWorkbook()` is the only public legacy ingress.

## RED evidence

- New controller tests initially failed 4 assertions because no schema 2 controller existed.
- The atomic `beforeNotify` rollback test then failed 1 assertion until document and legacy
  checkpoints were restored together.
- New React document-contract tests initially failed because the component still exposed
  `value`, `defaultValue`, `onChange`, and `getValue`.
- Consumer migration exposed expected RED failures in controlled reconciliation, Strict Mode
  cleanup ownership, TypeDoc/root exports, package fixtures, SSR fixtures, and browser
  canonical-shape assertions.

## GREEN implementation

- Added `SpreadsheetDocumentController`, with frozen document snapshots, revision/history
  delegation, stable schema sheet IDs, atomic dispatch/rollback, and subscription ownership.
- Centralized schema 2-to-legacy operation adaptation in
  `src/core/controller/runtime-projection.ts`; the engine consumes only the read-only
  `snapshot.projection`.
- Replaced React ingress with `document`, `defaultDocument`, and `onDocumentChange`.
- Replaced `TegoSheetHandle.getValue()` with frozen `getDocument()`.
- Preserved controlled optimistic acknowledgement, rollback, replacement, replay, callback
  ordering, history, editing, clipboard, validation, layout, and resource cleanup behavior.
- Removed root `WorkbookData` and `WorkbookInput` exports and migrated all public consumers,
  demos, package probes, examples, and handwritten documentation.

## Changed legacy contracts

- `value` -> `document`
- `defaultValue` -> `defaultDocument`
- `onChange` -> `onDocumentChange`
- `TegoSheetHandle.getValue()` -> `TegoSheetHandle.getDocument()`
- Root `WorkbookData` / `WorkbookInput` -> `SpreadsheetDocument`
- Direct sparse workbook ingress -> explicit `migrateLegacyWorkbook()` followed by schema 2
  ingress
- Default/empty legacy artifacts such as `freeze: "A1"`, empty autofilters, and duplicated
  cell-level merge metadata are omitted from schema 2 canonical snapshots.

## Verification

- `npx vitest run --project unit`: 39 files, 620 tests passed.
- `npx vitest run --project component`: 26 files, 202 tests passed.
- `npx vitest run --project architecture`: 14 files, 111 tests passed.
- `npm run test:ssr`: public ESM/CJS import probe and SSR component test passed.
- `npm run test:package`: 40 tests passed, including clean packed Vite, NodeNext ESM, and
  CommonJS consumers.
- `npm run typecheck`: passed.
- `npm run lint`: passed with warnings denied.
- `npm run format:check`: passed.
- `npm run build`: ESM/CJS bundles and declarations built.
- `npm run build:demo`: demo typecheck and production build passed.
- `npm run docs:build`: TypeDoc and Docusaurus client/server production builds passed.
- `npm run test:browser`: final six-project matrix passed 93 tests with 3 intentional
  desktop touch-only skips.

## Commits

- `3f6cbac feat(document): add schema 2 runtime controller`
- `d131eb5 feat(document): switch React ingress to schema 2`
- Final consumer/docs/verification commit: `feat(document): switch runtime to workbook 2`

## Remaining incompatibilities and concerns

- This is intentionally breaking: legacy React props, root workbook types, and `getValue()` no
  longer compile.
- Current editing operations still execute through the private legacy command engine projection.
  It is a single bounded adapter, not a second public or independently owned document truth.
- Existing React test-suite `act(...)` warnings in the read-only Suspense probe remain warnings;
  the suite passes and this task does not change that scheduling test.

## Review-fix pass: preserve schema 2 runtime truth

### RED evidence

- Direct schema 2 fixtures exposed five failures: constructor/dispatch lost custom input and
  schema-only cell references, validation references leaked across sheets, projection failure
  left split-brain state, rejected `beforeNotify` consumed a change sequence, and cloning
  dangerous JSON keys did not have an explicit safe contract.
- The first integrated component run then exposed nine projection regressions: controlled
  checkpoints captured the pre-commit document and validation-only commands retained the prior
  cell reference.
- The first fresh 96-case browser matrix exposed the same clear-format failure in all six browser
  projects because an absent operational `styleId` could not remove the previous value during
  object-spread merging.

### GREEN implementation

- Constructor and replace now preserve the parsed schema 2 document exactly; operation projection
  is one private boundary and never becomes document truth.
- Legacy-to-document projection compares the before/after operation boundary and merges only the
  fields the command changed. Custom input, resource/template references, metadata, settings,
  resources, templates, and extensions survive unrelated commands.
- Styles, validation references, editable/printable flags, and cell input are merged as
  operation-owned fields, so removal is represented correctly instead of reviving an omitted
  previous property.
- Validation collection is sheet-local and validation-only commands update references without
  creating cells on another sheet.
- Projection and user `beforeNotify` failures roll back legacy state, schema state, revision,
  history, subscriptions, and change sequencing as one transaction. Controlled reconciliation
  checkpoints capture the prepared schema document without publishing it early.
- Replace projects and validates before changing either runtime truth, so a rejected projection
  has no document-first visibility window.
- Deep frozen clones use null-prototype records and own-property definitions, preserving dangerous
  JSON keys without prototype mutation.
- `TegoSheetProps` is now a compile-time controlled/uncontrolled XOR. Controlled read-only
  consumers may omit `onDocumentChange`; editable controlled consumers are documented to apply
  emitted snapshots.
- Both packaged and website migration guides now use schema 2 props/handle methods and describe
  persisted `workbook.sheets[].id` identity correctly.

### Final verification

- Focused review regression tests: 12 controller tests and 4 document-contract component tests
  passed.
- `npx vitest run --project unit`: 39 files, 627 tests passed.
- `npx vitest run --project component`: 26 files, 203 tests passed.
- `npx vitest run --project architecture`: 14 files, 111 tests passed.
- `npm run typecheck`, `npm run lint`, `npm run format:check`, and `npm run build`: passed.
- `npm run test:ssr`: public ESM/CJS import probe plus SSR component test passed.
- `npm run test:package`: 40 tests passed, including packed ESM/CommonJS declarations and clean
  Vite consumer builds.
- `npm run test:browser`: fresh six-project matrix passed 93 tests with 3 intentional desktop
  touch-only skips.
- The parity manifest's 32 structural checks passed; its retained release-evidence check correctly
  refused to run against the dirty pre-commit worktree because its revision fingerprint requires a
  clean repository.
