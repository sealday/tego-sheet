# DATA-01 Outline Grouping and Enhanced Autofill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Persist safe worksheet outline groups and provide revision-bound typed/AST-aware
autofill.

**Architecture:** Workbook 2.0 remains the document truth. Group commands mutate only the schema
projection; collapsed visibility is derived at runtime/output boundaries. Autofill preview
publishes the existing public command while the schema handler performs typed pattern and formula
AST transforms.

**Tech Stack:** TypeScript 6, Vitest, immutable Workbook 2.0 documents, versioned controller
commands.

## Global Constraints

- Do not modify files under `src/template`.
- All mutations use versioned commands and one undoable transaction.
- Formula translation uses `parseFormula` + `translateFormula` + `renderFormula`.
- Long planning loops yield every 256 cells and honor `AbortSignal`.
- Group violations use `GROUP_LIMIT_EXCEEDED`.

---

### Task 1: Persistent outline model and parser

**Files:**
- Modify: `src/document/model/document.ts`
- Modify: `src/document/parse-document.ts`
- Test: `tests/unit/document/outline-groups.test.ts`

**Interfaces:**
- Produces: `SheetGroup`, `Sheet.groups`, canonical group parsing and stable diagnostics.

- [ ] **Step 1: Write failing parser tests**

Create fixtures with valid nested row/column groups and invalid duplicate IDs, crossing ranges,
out-of-bounds ranges, and excessive group counts. Assert canonical order/levels and stable
diagnostics.

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test -- --run tests/unit/document/outline-groups.test.ts`

Expected: FAIL because `groups` is not part of `Sheet`.

- [ ] **Step 3: Implement the minimal model/parser**

Add:

```ts
interface SheetGroup {
  readonly id: string;
  readonly axis: 'row' | 'column';
  readonly start: number;
  readonly end: number;
  readonly level: number;
  readonly collapsed: boolean;
}
```

Parse exact fields, reject invalid interval topology, normalize levels from containment, and
canonicalize by axis/start/end/id.

- [ ] **Step 4: Run the parser tests and commit**

Run the focused test plus `npm run typecheck`, then commit the model/parser slice.

### Task 2: Group commands, history, and structural transforms

**Files:**
- Modify: `src/core/commands/workbook-command.ts`
- Modify: `src/core/commands/validate-command.ts`
- Modify: `src/core/commands/schema-command-plan.ts`
- Modify: `src/core/controller/spreadsheet-document-controller.ts`
- Modify: `src/core/controller/runtime-projection.ts`
- Test: `tests/unit/core/outline-group-commands.test.ts`

**Interfaces:**
- Consumes: `SheetGroup`.
- Produces: `group`, `ungroup`, `toggle-group` commands and collapsed visibility derivation.

- [ ] **Step 1: Write failing command/history/structure tests**

Assert group/toggle/ungroup transactions, permission gates, undo/redo, insertion/deletion endpoint
updates, empty-group removal, level normalization, explicit-hidden union, and unchanged persisted
row/column hidden flags.

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test -- --run tests/unit/core/outline-group-commands.test.ts`

Expected: FAIL because the command discriminators are unknown.

- [ ] **Step 3: Implement commands and schema projection**

Validate exact IDs/ranges, mutate groups in `prepareSchemaCommand`, include the commands in
schema-only commits, transform group ranges during structural commands, and union collapsed
indexes into runtime row/column `hide`.

- [ ] **Step 4: Run focused tests and commit**

Run outline, controller, transaction, and typecheck suites; commit the command slice.

### Task 3: Typed and AST-aware autofill

**Files:**
- Modify: `src/core/commands/schema-command-plan.ts`
- Modify: `src/data-tools/transform-planner.ts`
- Test: `tests/unit/data-tools/enhanced-autofill.test.ts`

**Interfaces:**
- Produces: `AutofillTransform` preview request using the existing `DataTransformPreview`.

- [ ] **Step 1: Write failing preview and semantic tests**

Assert no mutation before commit, one transaction after commit, typed scalar copy, numeric/date/
suffix patterns, AST-relative and absolute formula references, stale plans, cell budgets, and
post-invocation cancellation.

- [ ] **Step 2: Run the test and verify RED**

Run: `npm test -- --run tests/unit/data-tools/enhanced-autofill.test.ts`

Expected: FAIL because `autofill` is not a data transform and schema autofill tiles inputs.

- [ ] **Step 3: Implement minimal planner and schema transform**

Snapshot the request, preview bounded target samples with chunked yielding, publish one
versioned `autofill` command, and map schema cells using typed pattern inference. Translate
formulas with:

```ts
renderFormula(translateFormula(parseFormula(source), { rowDelta, columnDelta }))
```

- [ ] **Step 4: Run focused tests and commit**

Run enhanced autofill, existing range/autofill, controller, typecheck, lint, and format checks;
commit the autofill slice.

### Task 4: Output visibility and final verification

**Files:**
- Modify only non-template runtime/output projection files proven necessary by RED tests.
- Test: relevant print/display/output suites.

**Interfaces:**
- Consumes: persisted groups.
- Produces: identical editor/print derived hidden semantics.

- [ ] **Step 1: Add failing output visibility tests**

Assert collapsed rows/columns are omitted while explicitly hidden rows remain hidden after group
expansion.

- [ ] **Step 2: Implement the smallest shared derived visibility helper**

Use the helper from runtime and output projections without mutating `Sheet.rows` or
`Sheet.columns`.

- [ ] **Step 3: Run full verification and commit**

Run focused suites, `npm run typecheck`, `npm run lint`, and `npm run format:check`; report any
unrelated worktree changes separately.
