# Complete Product Roadmap Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement every planned capability in the approved tego-sheet product Roadmap and move each item to a verified shipped record.

**Architecture:** Replace the current workbook-centric implementation with the compiler-centered `SpreadsheetDocument` architecture defined by the Mini-RFCs. Delivery follows the hard dependency path `F1 → F5 → F2 → F3 → F4 → TP1`; later work is released in dependency-safe waves, and every output adapter consumes the same immutable `GeneratedDocument`.

**Tech Stack:** TypeScript 6, React 19, Canvas 2D, SVG, Vitest, Testing Library, Playwright, Vite, Docusaurus, optional tree-shakeable output/interchange adapters selected by recorded dependency evaluations.

## Global Constraints

- Keep the product positioned as an embeddable React component and TypeScript SDK; do not add accounts, cloud storage, background jobs, or hosted collaboration services.
- Schema 2 `SpreadsheetDocument` is the only persistent source of truth; React, Canvas, DOM, selection, formula caches, adapter instances, and session state are never serialized into it.
- All document changes pass through Command/Transaction. UI, plugins, renderers, and adapters never mutate snapshots directly.
- The same explicit snapshot, locale, time zone, date system, clock, font metrics, data, resources, and print profile must produce deterministic calculation, presentation, pagination, and output geometry.
- Template expressions are interpreted by a restricted DSL and never execute JavaScript, access prototypes/globals, assign values, or invoke undeclared functions.
- Preview, browser print, PDF, SVG, PNG, and XLSX output consume one immutable `GeneratedDocument`; output adapters never access a mutable controller or re-layout pages.
- Every asynchronous stage accepts `AbortSignal` and enforces explicit cell, node, page, byte, decompression, time, and memory limits before exposing partial output.
- Excel/XLSX semantics are the compatibility baseline only for fixtures and functions explicitly listed as supported.
- Breaking the unpublished legacy API is allowed; migration fixtures and a migration guide replace parallel compatibility branches.
- No new runtime dependency is added until its maintenance, browser/Worker support, bundle size, license, security, CJK/font, deterministic-output, and tree-shaking constraints are recorded in a dependency decision.
- Each behavior follows RED → GREEN → REFACTOR. A task is not complete until its focused tests, the affected project tests, typecheck, lint, build, and architecture checks pass.
- Each completed Roadmap item moves from `planned` to the shipped record only after its Mini-RFC acceptance criteria and interoperability matrix pass.

## Delivery Map

| Wave | Roadmap coverage | Hard prerequisites | Parallel lanes after prerequisite |
| --- | --- | --- | --- |
| 0A | F1 Workbook 2.0 | none | schema, migration fixtures, controller ingress audit |
| 0B | F5 Registry Kernel | F1 schema primitives | cell types, lifecycle tests |
| 0C | F2 Transactions | F1, F5 boundaries | command handlers, coordinate transforms |
| 0D | F3 Formula/Format | F1, F2 | formula graph, formatter matrix |
| 0E | F4 Presentation | F1–F3 | Canvas adapter, accessibility viewport, display list |
| 1 | TP1 Template Print MVP | Wave 0 | compiler, layout, designer UI, print adapter |
| 2A | TP2, TP3 | TP1 | advanced structures and resource pipeline |
| 2B | TP4, TP5, TP6 | TP3 plus adapter evaluations | PDF, XLSX, image adapters |
| 3A | FMT, VAL, FRM, VIEW, DATA, IO, OBJ | Wave 0; TP3 for objects | seven independent capability lanes |
| 3B | TBL, CHT, SPK, PVT | 3A dependencies | chart/sparkline and pivot lanes |
| 3C | SLC, GSK, SLV | 3B dependencies for slicer | slicer and solving lanes |
| 4A | E3, E1, E2 | Waves 0–3 | cell and template SDKs after registry |
| 4B | H3, H1, H4, H5, H2, H6 | E3 and transaction contracts | comments/history overlap after persistence revision |

---

### Task 1: Establish the executable acceptance ledger

**Files:**
- Create: `docs/roadmap-implementation/acceptance.ts`
- Create: `tests/unit/roadmap/acceptance.test.ts`
- Modify: `website/src/data/roadmap.ts`
- Create: `website/docs/roadmap/shipped.md`

**Interfaces:**
- Consumes: the 33 canonical item IDs and phase assignments from `website/src/data/roadmap.ts`.
- Produces: `roadmapAcceptance`, `getRoadmapDeliveryState()`, and one machine-checked mapping from every planned item to its Mini-RFC section, implementation tasks, and verification commands.

- [ ] **Step 1: Write the failing coverage test**

```ts
import { roadmapItems } from '../../../website/src/data/roadmap';
import { roadmapAcceptance } from '../../../docs/roadmap-implementation/acceptance';

it('maps every roadmap item to acceptance evidence without duplicate ids', () => {
  expect(new Set(roadmapAcceptance.map((entry) => entry.id)).size).toBe(33);
  expect(roadmapAcceptance.map((entry) => entry.id).sort()).toEqual(
    roadmapItems.map((entry) => entry.id).sort(),
  );
  expect(roadmapAcceptance.every((entry) => entry.spec && entry.tests.length > 0)).toBe(true);
});
```

- [ ] **Step 2: Run the focused test and verify RED**

Run: `npx vitest run --project unit tests/unit/roadmap/acceptance.test.ts`

Expected: FAIL because `docs/roadmap-implementation/acceptance.ts` does not exist.

- [ ] **Step 3: Add stable item IDs and the typed acceptance ledger**

```ts
export interface RoadmapAcceptanceEntry {
  readonly id: string;
  readonly phase: 0 | 1 | 2 | 3 | 4;
  readonly spec: `website/docs/roadmap/${string}.md#${string}`;
  readonly tasks: readonly string[];
  readonly tests: readonly string[];
  readonly state: 'planned' | 'in-progress' | 'shipped';
}
```

Populate all 33 entries from `website/docs/roadmap/index.md`; keep each entry `planned` until its evidence is produced.

- [ ] **Step 4: Run focused Roadmap tests and verify GREEN**

Run: `npx vitest run --project unit tests/unit/roadmap/acceptance.test.ts tests/unit/website/roadmap.test.ts`

Expected: PASS with 33 unique mapped items.

- [ ] **Step 5: Commit the acceptance ledger**

```bash
git add docs/roadmap-implementation website/src/data/roadmap.ts website/docs/roadmap/shipped.md tests/unit/roadmap
git commit -m "test(roadmap): track implementation acceptance"
```

### Task 2: Deliver F1 Workbook 2.0

**Files:**
- Create: `src/document/model/ids.ts`
- Create: `src/document/model/document.ts`
- Create: `src/document/model/sparse-cells.ts`
- Create: `src/document/diagnostics.ts`
- Create: `src/document/create-document.ts`
- Create: `src/document/parse-document.ts`
- Create: `src/document/serialize-document.ts`
- Create: `src/document/migrate-legacy.ts`
- Create: `tests/unit/document/document-roundtrip.test.ts`
- Create: `tests/unit/document/document-validation.test.ts`
- Create: `tests/unit/document/legacy-migration.test.ts`
- Modify: `src/index.ts`

**Interfaces:**
- Produces: `SpreadsheetDocument`, typed opaque IDs, `CellInput`, `Diagnostic`, `createSpreadsheetDocument()`, `parseSpreadsheetDocument()`, `serializeSpreadsheetDocument()`, and `migrateLegacyWorkbook()`.
- Preserves: stable sheet identity and explicit `excel-1900 | excel-1904` date systems.

- [ ] **Step 1: Write failing schema, type-preservation, deterministic-byte, deep-readonly, invalid-reference, conflict-merge, limit, and migration fixture tests**

```ts
const inputs: CellInput[] = [
  { type: 'blank' },
  { type: 'string', value: '' },
  { type: 'number', value: 45205 },
  { type: 'boolean', value: false },
  { type: 'formula', source: '=A1+1' },
];
const encoded = serializeSpreadsheetDocument(documentWithInputs(inputs));
expect(serializeSpreadsheetDocument(parseOk(encoded))).toBe(encoded);
expect(readInputs(parseOk(encoded))).toEqual(inputs);
```

- [ ] **Step 2: Run the document tests and verify RED**

Run: `npx vitest run --project unit tests/unit/document`

Expected: FAIL on missing Workbook 2.0 exports.

- [ ] **Step 3: Implement the schema 2 decoder, canonicalizer, cross-reference validator, structural-sharing snapshot, deterministic serializer, and legacy migration**

Use the exact public shapes from `website/docs/roadmap/foundation.md#f1-workbook-20`. Parser errors return aggregated diagnostics and never expose a partial document. Serializer preserves user-ordered arrays and sorts registries, IDs, and sparse numeric coordinates deterministically.

- [ ] **Step 4: Switch public construction and controller ingress/egress to Workbook 2.0**

Remove mutable reference sharing at props, callbacks, snapshots, and imperative-handle boundaries. Retain the legacy model only inside the pure one-shot migrator.

- [ ] **Step 5: Run F1 verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/document tests/unit/core/schema-validation.test.ts tests/unit/core/serialization.test.ts && npm run typecheck && npm run test:ssr`

Expected: PASS; invalid documents fail atomically with stable codes and paths.

- [ ] **Step 6: Commit F1**

```bash
git add src/document src/index.ts tests/unit/document tests/unit/core
git commit -m "feat(document): introduce workbook 2 schema"
```

### Task 3: Deliver F5 Minimal Extension and Adapter Kernel

**Files:**
- Create: `src/extensions/kernel/capabilities.ts`
- Create: `src/extensions/kernel/registry.ts`
- Create: `src/extensions/kernel/manifest.ts`
- Create: `src/extensions/cell-types/checkbox.ts`
- Create: `src/extensions/cell-types/dropdown.ts`
- Create: `tests/unit/extensions/registry.test.ts`
- Create: `tests/unit/extensions/cell-types.test.ts`
- Create: `tests/types/kernel-capabilities.test-d.ts`

**Interfaces:**
- Produces: internal `AdapterRegistryKernel`, declaration-mergeable `KernelCapabilities`, built-in `BuiltInCellTypeDefinition`, stable resolution, and idempotent disposal.
- Consumes: F1 `JsonValue`, custom cell input, diagnostics, immutable snapshots, and opaque IDs.

- [ ] **Step 1: Write failing compile-time and runtime tests**

```ts
await registry.register(checkboxRegistration);
expect(registry.resolve('cell-type', { id: 'checkbox', environment: 'browser' }))
  .toBe(checkboxRegistration.implementation);
await expect(registerDuplicate()).rejects.toMatchObject({ code: 'EXTENSION_DUPLICATE_ID' });
```

Cover exact selection, ambiguity, environment mismatch, incompatible major/minor, failed initialization compensation, unregister, repeated dispose, and multi-error cleanup.

- [ ] **Step 2: Run kernel tests and verify RED**

Run: `npx vitest run --project unit tests/unit/extensions`

Expected: FAIL because the kernel does not exist.

- [ ] **Step 3: Implement the typed registry and built-in cell definitions**

Implement atomic registration, stable `kind/id` listing, O(1) exact lookup, environment filtering, package-supported API version checks, and lifecycle cleanup. Checkbox/dropdown definitions own validation, migration, formatted value semantics, accessibility semantics, and formula scalar coercion.

- [ ] **Step 4: Run F5 verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/extensions && npm run typecheck && npm run lint`

Expected: PASS with no listener, task, or object-URL leaks.

- [ ] **Step 5: Commit F5**

```bash
git add src/extensions tests/unit/extensions tests/types
git commit -m "feat(extensions): add typed adapter kernel"
```

### Task 4: Deliver F2 Atomic Command and Transaction

**Files:**
- Create: `src/document/commands/command.ts`
- Create: `src/document/commands/patch.ts`
- Create: `src/document/commands/handlers.ts`
- Create: `src/document/commands/coordinate-transform.ts`
- Create: `src/document/controller/document-controller.ts`
- Create: `src/document/controller/transaction.ts`
- Create: `tests/unit/document/transactions.test.ts`
- Create: `tests/unit/document/coordinate-transform.test.ts`
- Create: `tests/property/document/transaction-invariants.test.ts`
- Modify: `src/react/hooks/use-controlled-workbook.ts`

**Interfaces:**
- Produces: `DocumentController.execute/transact/dryRun/undo/redo/snapshot/subscribe`, versioned serializable commands, internal-only patches, inverse patches, revisions, one-event commits, and permission-gate hooks.
- Transforms: cells, formula AST references, names, merges, validations, conditional formats, templates, print areas, tables, and object anchors through one `CoordinateTransform`.

- [ ] **Step 1: Write failing atomicity, inverse, one-event, revision-conflict, dry-run, serialization, permission, and coordinate-transform tests**

```ts
const before = serializeSpreadsheetDocument(controller.snapshot());
const result = controller.transact([validCommand, invalidCommand]);
expect(result.committed).toBe(false);
expect(serializeSpreadsheetDocument(controller.snapshot())).toBe(before);
expect(listener).not.toHaveBeenCalled();
```

- [ ] **Step 2: Run transaction tests and verify RED**

Run: `npx vitest run --project unit tests/unit/document/transactions.test.ts tests/unit/document/coordinate-transform.test.ts`

Expected: FAIL because the document controller and coordinate transform are missing.

- [ ] **Step 3: Implement validate → authorize → normalize → plan → transform → invariant-check → atomic-commit**

Patch construction remains private to handlers. Undo applies minimal inverse patches as a new revision; redo reapplies semantic patches. Failed transactions retain byte-identical state and revision.

- [ ] **Step 4: Migrate every current edit path to commands and delete direct mutable writes**

Replace old controller mutation calls in core operations, React hooks, interaction adapters, toolbar, dialogs, clipboard, sort/filter, and structural editing. Session selection/scroll/dialog state remains outside document history.

- [ ] **Step 5: Run F2 verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/document tests/unit/core tests/unit/react tests/property/document && npm run typecheck && npm run lint`

Expected: PASS; every batch emits one committed event and one undo entry.

- [ ] **Step 6: Commit F2**

```bash
git add src/document src/core src/react src/engine src/ui tests
git commit -m "feat(commands): make document changes transactional"
```

### Task 5: Deliver F3 Formula Dependency and Number-Format Core

**Files:**
- Create: `src/formula/ast.ts`
- Create: `src/formula/parser.ts`
- Create: `src/formula/reference-resolver.ts`
- Create: `src/formula/dependency-graph.ts`
- Create: `src/formula/evaluator.ts`
- Create: `src/formula/function-registry.ts`
- Create: `src/format/number-format-parser.ts`
- Create: `src/format/number-formatter.ts`
- Create: `tests/unit/formula/parser.test.ts`
- Create: `tests/unit/formula/dependency-graph.test.ts`
- Create: `tests/unit/formula/excel-compatibility.test.ts`
- Create: `tests/unit/format/number-format.test.ts`

**Interfaces:**
- Produces: typed `FormulaValue`, `FormulaProgram`, `FormulaEngine`, `NumberFormatter`, fixed `CalculationEnvironment`, stable standard errors, and explicit compatibility manifests.
- Consumes: F1 typed cells, stable IDs, F2 coordinate transforms, and F5 formula-function registration.

- [ ] **Step 1: Write failing parser, reference, error propagation, cycle, incremental-dirty, fixed-clock, function-matrix, date-system, locale, time-zone, and custom-format tests**

```ts
const result = engine.recalculate(program, [change('Sheet1!A1')], fixedEnvironment);
expect(result.evaluatedAddresses).toEqual(['Sheet1!B1', 'Sheet1!C1']);
expect(result.evaluatedAddresses).not.toContain('Sheet1!Z1');
```

- [ ] **Step 2: Run formula/format tests and verify RED**

Run: `npx vitest run --project unit tests/unit/formula tests/unit/format`

Expected: FAIL on missing typed formula and formatter modules.

- [ ] **Step 3: Implement lexer/parser, stable-reference AST, dependency graph, evaluator, registry, and formatter**

Never use `eval`, `Function`, prototype traversal, implicit system clock, implicit locale, or implicit time zone. Store source only in the document; keep AST, dependencies, and values in rebuildable `FormulaProgram`.

- [ ] **Step 4: Integrate structure transforms and remove legacy formula/cache truth**

Copy/insert/delete/template expansion transform AST nodes, reserialize normalized source, rebuild affected graph nodes, and recalculate once.

- [ ] **Step 5: Run F3 verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/formula tests/unit/format tests/unit/core/formulas.test.ts tests/unit/core/structure.test.ts && npm run typecheck && npm run lint`

Expected: PASS with deterministic typed values and formatted text.

- [ ] **Step 6: Commit F3**

```bash
git add src/formula src/format src/core tests/unit/formula tests/unit/format
git commit -m "feat(formula): add incremental typed calculation core"
```

### Task 6: Deliver F4 Shared Presentation, Accessibility, and Print Display List

**Files:**
- Create: `src/presentation/cell-presentation.ts`
- Create: `src/presentation/presentation-resolver.ts`
- Create: `src/presentation/presentation-cache.ts`
- Create: `src/presentation/font-metrics.ts`
- Create: `src/print/display-list.ts`
- Create: `src/react/accessibility/accessibility-grid.tsx`
- Modify: `src/engine/canvas/cell-painter.ts`
- Modify: `src/engine/canvas/canvas-engine.ts`
- Create: `tests/unit/presentation/presentation.test.ts`
- Create: `tests/component/accessibility-grid.test.tsx`
- Create: `tests/architecture/renderer-boundaries.test.ts`

**Interfaces:**
- Produces: `CellPresentation`, `ResolvedStyle`, `PresentationResolver`, deterministic `PrintDisplayList`, injected font metrics, and viewport-bounded semantic DOM.
- Consumes: immutable document snapshot, formula program, revisions, explicit presentation environment, and read-only geometry.

- [ ] **Step 1: Write failing cross-surface parity, deterministic display-list, viewport-size DOM, focus/edit/selection/error semantics, cache budget, missing-font, and architecture-boundary tests**

```ts
expect(canvasPresentation.formattedText).toBe(accessibleCell.textContent);
expect(printCell.text).toBe(canvasPresentation.formattedText);
expect(accessibilityContainer.querySelectorAll('[role="gridcell"]').length).toBeLessThan(500);
```

- [ ] **Step 2: Run presentation tests and verify RED**

Run: `npx vitest run --project unit --project component --project architecture tests/unit/presentation tests/component/accessibility-grid.test.tsx tests/architecture/renderer-boundaries.test.ts`

Expected: FAIL because Canvas still owns part of formatting/printing and the accessibility grid is missing.

- [ ] **Step 3: Implement presentation resolution and migrate Canvas to read-only batches**

Resolve value, formatted text, style, validation, annotations, and visibility once. Cache by document, calculation, condition, style, and environment revisions with explicit entry/byte budgets.

- [ ] **Step 4: Implement the virtual accessibility grid and device-independent display list**

The accessibility layer renders viewport plus bounded overscan, retains a single active gridcell, summarizes large selections, and returns focus after editor close. Print layout never reads selection, scroll, zoom, DPR, toolbar, dialogs, or Canvas bitmap state.

- [ ] **Step 5: Run F4 verification and verify GREEN**

Run: `npx vitest run --project unit --project component --project architecture tests/unit/presentation tests/unit/engine tests/component/accessibility-grid.test.tsx tests/architecture && npm run test:browser && npm run typecheck`

Expected: PASS with renderer mutation imports rejected.

- [ ] **Step 6: Commit F4**

```bash
git add src/presentation src/print src/engine src/react/accessibility tests
git commit -m "feat(rendering): unify cell presentation semantics"
```

### Task 7: Deliver TP1 Template Print MVP

**Files:**
- Create: `src/template/model.ts`
- Create: `src/template/expression/*`
- Create: `src/template/compiler/*`
- Create: `src/template/expand/*`
- Create: `src/print/layout/*`
- Create: `src/output/browser-print-adapter.ts`
- Create: `src/react/template-designer/*`
- Create: `src/react/preview/*`
- Create: `tests/unit/template/*`
- Create: `tests/unit/print/*`
- Create: `tests/component/template-designer.test.tsx`
- Create: `tests/browser/template-print.spec.ts`
- Delete: `src/engine/canvas/print-renderer.ts`
- Delete: `src/ui/print-workbook.ts`

**Interfaces:**
- Produces: `SpreadsheetTemplate`, `CompiledTemplate`, `GeneratedDocument`, restricted expression IR, scalar/repeat-row/conditional expansion, deterministic pagination, SVG preview, and isolated iframe print.
- Enforces: `CompiledTemplate.sourceDocumentHash === RenderRequest.currentDocumentHash`.

- [ ] **Step 1: Write failing compiler, DSL safety, stale-source, repeat/conditional transform, target, pagination, preview parity, iframe isolation, afterprint, timeout, cancellation, and cleanup tests**

```ts
expect(() => compileExpression('value.constructor.constructor("return globalThis")()'))
  .toThrowErrorMatchingObject({ code: 'TEMPLATE_EXPRESSION_UNSAFE' });
expect(renderWithHash(compiled, changedDocumentHash)).toMatchObject({
  diagnostics: [expect.objectContaining({ code: 'TEMPLATE_SOURCE_STALE' })],
});
```

- [ ] **Step 2: Run TP1 tests and verify RED**

Run: `npx vitest run --project unit --project component tests/unit/template tests/unit/print tests/component/template-designer.test.tsx`

Expected: FAIL because the compiler-centered pipeline does not exist.

- [ ] **Step 3: Implement compiler, safe resolver, expansion, recalc, presentation, and pagination pipeline**

Compile bindings from metadata, not cell-string scanning. Expand on an isolated workbook, transform all range metadata and relative formula AST references, recalculate affected nodes once, then paginate using explicit paper, margin, scale, repeated-title, break, locale, time-zone, and font inputs.

- [ ] **Step 4: Implement Template/Preview modes and isolated browser printing**

Template mode edits the same model used by the SDK. Preview renders SVG from `PrintDisplayList`. Browser print writes only generated pages into an isolated iframe and cleans it after `afterprint`, cancellation, or a bounded timeout.

- [ ] **Step 5: Delete the legacy viewport/Canvas print path and run TP1 verification**

Run: `npx vitest run --project unit --project component --project architecture tests/unit/template tests/unit/print tests/component/template-designer.test.tsx tests/architecture && npm run test:browser && npm run test:visual && npm run typecheck && npm run lint`

Expected: PASS; preview and print have identical page count and geometry.

- [ ] **Step 6: Commit TP1**

```bash
git add src/template src/print src/output src/react src/engine src/ui tests
git commit -m "feat(template): deliver deterministic print workflow"
```

### Task 8: Deliver TP2 Advanced Template Structures and TP3 Resource Pipeline

**Files:**
- Create: `src/template/compiler/region-tree.ts`
- Create: `src/template/expand/advanced-repeat.ts`
- Create: `src/template/compiler/subtemplates.ts`
- Create: `src/resources/resource-store.ts`
- Create: `src/resources/resource-session.ts`
- Create: `src/resources/image-decoder.ts`
- Create: `src/resources/font-loader.ts`
- Create: `src/resources/qr.ts`
- Create: `tests/unit/template/advanced-structures.test.ts`
- Create: `tests/unit/resources/resource-pipeline.test.ts`

**Interfaces:**
- Produces: nested/horizontal/two-dimensional/per-page/per-sheet repeats, subtemplate dependency validation, structural mapping, content-hash resource dedupe, explicit host resolver, font readiness, QR resources, and resource-session disposal.
- Consumes: TP1 render sessions, F2 transforms, F5 resource capability, and `AbortSignal`.

- [ ] **Step 1: Write failing nested scope, patch-order, circular-subtemplate, object-copy-policy, dedupe, quota, unsafe-SVG, no-network, font-wait, cancellation, and disposal tests**
- [ ] **Step 2: Run focused tests and verify RED**

Run: `npx vitest run --project unit tests/unit/template/advanced-structures.test.ts tests/unit/resources/resource-pipeline.test.ts`

Expected: FAIL on missing advanced structures and resource session.

- [ ] **Step 3: Implement region containment, outer-to-inner resolution, bottom-right-to-top-left patches, and structural maps**
- [ ] **Step 4: Implement resource resolution with hash dedupe, byte/decompression limits, safe decoding, explicit network policy, and idempotent cleanup**
- [ ] **Step 5: Run focused and TP1 regression tests and verify GREEN**

Run: `npx vitest run --project unit tests/unit/template tests/unit/resources tests/unit/print && npm run typecheck && npm run lint`

Expected: PASS with atomic failure before partial generated output.

- [ ] **Step 6: Commit TP2/TP3**

```bash
git add src/template src/resources tests/unit/template tests/unit/resources
git commit -m "feat(template): add advanced structures and resources"
```

### Task 9: Deliver TP4 PDF, TP5 XLSX, and TP6 Image Output

**Files:**
- Create: `docs/decisions/pdf-adapter.md`
- Create: `docs/decisions/xlsx-adapter.md`
- Create: `src/output/pdf/*`
- Create: `src/output/xlsx/*`
- Create: `src/output/image/*`
- Create: `tests/unit/output/pdf.test.ts`
- Create: `tests/unit/output/xlsx.test.ts`
- Create: `tests/unit/output/image.test.ts`
- Create: `tests/fixtures/output/*`

**Interfaces:**
- Produces: optional tree-shakeable PDF Blob, XLSX Blob, SVG pages, and PNG pages from one `GeneratedDocument`.
- Preserves: page geometry, typed cells, formulas, styles, merges, validation, conditional formatting, print settings, objects, embedded resources, and structured unsupported-feature diagnostics.

- [ ] **Step 1: Record reproducible dependency decisions before adding runtime packages**
- [ ] **Step 2: Write failing golden geometry, text-search, CJK/font, OOXML structure, deterministic ZIP, unsupported-feature, SVG safety, PNG DPI, multi-page, abort, and no-partial-output tests**
- [ ] **Step 3: Run output tests and verify RED**

Run: `npx vitest run --project unit tests/unit/output`

Expected: FAIL because output adapters are missing.

- [ ] **Step 4: Implement each adapter as a pure translation of `GeneratedDocument`**

PDF and image adapters translate `PrintDisplayList` without re-layout. XLSX serializes `RenderedWorkbook` and print metadata; stable package ordering is deterministic.

- [ ] **Step 5: Run structural, visual, package, browser, Worker, and office interoperability verification**

Run: `npx vitest run --project unit tests/unit/output && npm run test:browser && npm run test:visual && npm run test:package && npm run build`

Expected: PASS; recorded manual matrix confirms Excel Desktop, Excel for web, LibreOffice, Chrome, Firefox, Safari, and one Worker environment.

- [ ] **Step 6: Commit Phase 2 outputs**

```bash
git add docs/decisions src/output tests
git commit -m "feat(output): add pdf xlsx and image adapters"
```

### Task 10: Deliver Phase 3 Formula, Formatting, Validation, View, Data, IO, and Object Foundations

**Files:**
- Create: `src/format/conditional/*`
- Create: `src/validation/*`
- Extend: `src/formula/*`
- Create: `src/views/*`
- Create: `src/data-tools/*`
- Create: `src/interchange/*`
- Create: `src/objects/*`
- Create: `tests/unit/format/conditional-format.test.ts`
- Create: `tests/unit/validation/*`
- Create: `tests/unit/formula/advanced-formulas.test.ts`
- Create: `tests/unit/views/*`
- Create: `tests/unit/data-tools/*`
- Create: `tests/unit/interchange/*`
- Create: `tests/unit/objects/*`

**Interfaces:**
- Produces: conditional formatting, advanced validation/dropdown/checkbox, cross-sheet/named/structured/array formulas, saved views, transactional cleanup commands, CSV/TSV/XLSX/ODS adapters, and anchored images/shapes/text boxes.
- Shares: typed values, Formula dependency graph, F2 transactions, F5 capability injection, `CellPresentation`, resources, and print display lists.

- [ ] **Step 1: Write one failing acceptance suite per capability before its implementation lane begins**
- [ ] **Step 2: Run all Phase 3A tests and verify RED for each missing capability**

Run: `npx vitest run --project unit tests/unit/format tests/unit/validation tests/unit/formula/advanced-formulas.test.ts tests/unit/views tests/unit/data-tools tests/unit/interchange tests/unit/objects`

Expected: each new suite fails for its missing public behavior.

- [ ] **Step 3: Implement the seven lanes in dependency-safe parallel tasks**

Every document change is one validated transaction. Views keep derived row visibility outside persistent row truth. Data-tool previews bind to a base revision and reject stale commits. Readers parse atomically with resource limits. Object anchors transform through F2.

- [ ] **Step 4: Run Phase 3A verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/format tests/unit/validation tests/unit/formula tests/unit/views tests/unit/data-tools tests/unit/interchange tests/unit/objects && npm run test:browser && npm run typecheck && npm run lint`

Expected: PASS with output and XLSX parity fixtures.

- [ ] **Step 5: Commit Phase 3A**

```bash
git add src tests
git commit -m "feat(spreadsheet): deepen formulas data and objects"
```

### Task 11: Deliver Tables, Charts, Sparklines, Pivot, Slicer, Goal Seek, and Solver

**Files:**
- Create: `src/analysis/tables/*`
- Create: `src/analysis/charts/*`
- Create: `src/analysis/sparklines/*`
- Create: `src/analysis/pivot/*`
- Create: `src/analysis/slicer/*`
- Create: `src/analysis/goal-seek/*`
- Create: `src/analysis/solver/*`
- Create: `docs/decisions/solver-adapter.md`
- Create: `tests/unit/analysis/*`
- Create: `tests/component/analysis/*`

**Interfaces:**
- Produces: stable-ID structured tables/references, normalized charts, sparklines, PivotTable snapshots, slicer filter contributions, isolated incremental goal seek, and optional solver adapters.
- Consumes: Phase 3A formula/view/object/resource/transaction capabilities and F5 solver/chart registrations.

- [ ] **Step 1: Write failing acceptance tests for each analysis object, dependency invalidation, output parity, cancellation, budget, undo, and unsupported degradation**
- [ ] **Step 2: Run analysis tests and verify RED**

Run: `npx vitest run --project unit --project component tests/unit/analysis tests/component/analysis`

Expected: FAIL because analysis modules are missing.

- [ ] **Step 3: Implement Tables first, then Chart/Sparkline/Pivot, then Slicer**
- [ ] **Step 4: Implement Goal Seek in an isolated calculation context and Solver behind the evaluated adapter contract**
- [ ] **Step 5: Run Phase 3B/3C verification and verify GREEN**

Run: `npx vitest run --project unit --project component tests/unit/analysis tests/component/analysis && npm run test:browser && npm run test:visual && npm run typecheck && npm run lint`

Expected: PASS; Goal Seek recalculates only the variable cell's transitive dependents and Solver cannot access the controller.

- [ ] **Step 6: Commit Phase 3 analysis**

```bash
git add src/analysis docs/decisions tests/unit/analysis tests/component/analysis
git commit -m "feat(analysis): add tables visualization and solving"
```

### Task 12: Deliver Public Cell, Template Module, and Adapter SDKs

**Files:**
- Create: `src/sdk/adapters/*`
- Create: `src/sdk/cells/*`
- Create: `src/sdk/templates/*`
- Create: `src/sdk/trust/*`
- Create: `tests/types/public-sdk.test-d.ts`
- Create: `tests/unit/sdk/*`
- Create: `tests/worker/sdk-isolation.test.ts`

**Interfaces:**
- Produces: E3 public adapter lifecycle/capability/trust policy, E1 structured cell renderer/editor SDK, and E2 versioned template-module stage SDK.
- Extends: F5 rather than creating a parallel registry.

- [ ] **Step 1: Write failing type, manifest negotiation, scope, budget, cancellation, disposal, unknown-cell degradation, worker isolation, module ownership, deterministic IR/draw, and missing-module tests**
- [ ] **Step 2: Run SDK tests and verify RED**

Run: `npx vitest run --project unit tests/unit/sdk`

Expected: FAIL because public extension protocols do not exist.

- [ ] **Step 3: Publicize the registry contract with explicit `trusted-main | isolated-worker` execution**
- [ ] **Step 4: Implement Cell and Template Module SDKs on that runtime**

Worker plugins receive schema-cloned capability-specific inputs and never DOM, Canvas, controller, globals, implicit network, or undeclared resources. Trusted-main is explicitly not described as a malicious-code sandbox.

- [ ] **Step 5: Run SDK verification and verify GREEN**

Run: `npx vitest run --project unit tests/unit/sdk && npm run typecheck && npm run test:package && npm run build`

Expected: PASS with tree-shakeable optional adapters and no leaked resources.

- [ ] **Step 6: Commit public SDKs**

```bash
git add src/sdk tests
git commit -m "feat(sdk): publish cell template and adapter protocols"
```

### Task 13: Deliver Host Integration Adapter Contracts

**Files:**
- Create: `src/integrations/permission/*`
- Create: `src/integrations/persistence/*`
- Create: `src/integrations/comments/*`
- Create: `src/integrations/history/*`
- Create: `src/integrations/collaboration/*`
- Create: `src/integrations/ai/*`
- Create: `tests/contract/integrations/*`
- Create: `tests/component/integrations/*`

**Interfaces:**
- Produces: permission action/target/default-deny, revision persistence, durable comment outbox, semantic history/recovery, remote transaction/presence synchronization, and validated AI command proposals.
- Does not produce: a backend, account system, cloud store, model provider, API-key manager, or server authorization substitute.

- [ ] **Step 1: Write failing adapter contract suites in required order: Permission → Persistence → Comments → Version History → Collaboration → AI**
- [ ] **Step 2: Run integration contract tests and verify RED**

Run: `npx vitest run --project unit --project component tests/contract/integrations tests/component/integrations`

Expected: FAIL because host adapter contracts and UI connection points are missing.

- [ ] **Step 3: Implement permission and persistence foundations**

The same permission path gates toolbar, shortcuts, context menu, public command API, comments, output, and AI proposal. Server authorization remains mandatory. Persistence distinguishes pending local revision, acknowledged revision, retry, conflict, and disposal.

- [ ] **Step 4: Implement comments/history, then collaboration, then AI proposals**

Comments transform anchors through structural patches and use an idempotent outbox. Recovery creates a new version. Collaboration validates duplicate, out-of-order, stale, and illegal remote transactions atomically; presence is session-only. AI receives minimized/redacted context, returns schema-whitelisted proposals, passes permission and dry-run, expires on revision/permission change, and requires explicit user confirmation.

- [ ] **Step 5: Run host-integration verification and verify GREEN**

Run: `npx vitest run --project unit --project component tests/contract/integrations tests/component/integrations && npm run test:browser && npm run typecheck && npm run lint`

Expected: PASS, including the complete core suite with no adapter installed.

- [ ] **Step 6: Commit host integrations**

```bash
git add src/integrations tests
git commit -m "feat(integrations): add host adapter contracts"
```

### Task 14: Complete migration, documentation, examples, and Roadmap state

**Files:**
- Rewrite: `docs/migration-from-x-data-spreadsheet.md`
- Create: `docs/migration-to-workbook-2.md`
- Create: `website/docs/api/templates.md`
- Create: `website/docs/api/output.md`
- Create: `website/docs/api/extensions.md`
- Create: `website/docs/api/integrations.md`
- Modify: `website/src/data/roadmap.ts`
- Modify: `website/docs/roadmap/index.md`
- Modify: `website/docs/roadmap/shipped.md`
- Modify: `readme.md`
- Modify: `demo/*`

**Interfaces:**
- Produces: one-shot migration guidance, public API examples, adapter responsibility boundaries, supported compatibility matrices, and shipped evidence for all 33 Roadmap items.

- [ ] **Step 1: Write failing docs/example/link/Roadmap-state tests**
- [ ] **Step 2: Run docs tests and verify RED**

Run: `npm run typecheck:docs && npm run test:docs`

Expected: FAIL until new API docs, demo, and shipped links exist.

- [ ] **Step 3: Document exact removed APIs, migration functions, diagnostics, limits, compatibility matrices, and adapter examples**
- [ ] **Step 4: Move only acceptance-ledger entries with verified evidence from planned to shipped**
- [ ] **Step 5: Build and test documentation and verify GREEN**

Run: `npm run docs:build && npm run test:docs && npm run test:docs-visual`

Expected: PASS with no broken links, stale planned items, console errors, or horizontal overflow.

- [ ] **Step 6: Commit documentation and Roadmap state**

```bash
git add docs website readme.md demo tests
git commit -m "docs: publish completed product roadmap"
```

### Task 15: Run final interoperability, architecture, security, and release gates

**Files:**
- Create: `docs/roadmap-implementation/final-verification.md`
- Update: `docs/roadmap-implementation/acceptance.ts`

**Interfaces:**
- Consumes: all task evidence and all Mini-RFC acceptance criteria.
- Produces: a complete verification matrix with command output, browser/office/manual evidence, supported/degraded feature declarations, and independent review results.

- [ ] **Step 1: Run the full automated gate from a clean checkout**

```bash
npm run format:check
npm run lint
npm run typecheck
npm run test
npm run test:browser
npm run test:visual
npm run test:ssr
npm run test:package
npm run build
npm run build:demo
npm run docs:build
npm run typecheck:docs
npm run test:docs
npm run test:docs-visual
```

Expected: every command exits 0 with no warnings treated as ignored failures.

- [ ] **Step 2: Run architecture and security audits**

Prove renderers/adapters cannot import mutation capabilities; template expressions cannot escape; external resources are opt-in; SVG/ZIP/XML/CSV inputs enforce quotas; Worker plugins are capability-scoped; permissions default deny; AI logs contain no cell values, instruction, model output, or API key.

- [ ] **Step 3: Run the interoperability matrix**

Verify browser print in current stable Chrome, Firefox, and Safari; PDF/image generation in those browsers and one Worker; XLSX in Excel Desktop, Excel for web, and LibreOffice; ODS in LibreOffice; accessibility with keyboard plus at least VoiceOver and one additional screen reader/browser pair.

- [ ] **Step 4: Reconcile all 33 acceptance-ledger entries against real evidence**

No entry becomes `shipped` from test inference alone when its Mini-RFC requires manual browser, Office, accessibility, license, security, or visual evidence.

- [ ] **Step 5: Run changed-file cleanup, repeat the full gate, and request independent code and architecture review**

Apply `ai-slop-cleaner` only to changed files, rerun Step 1, then require code-review recommendation `APPROVE`, architecture status `CLEAR`, and proof for every non-negotiable invariant in the approved design.

- [ ] **Step 6: Commit final verification**

```bash
git add docs/roadmap-implementation website/src/data/roadmap.ts
git commit -m "chore(release): verify complete product roadmap"
```

## Plan Self-Review

- **Spec coverage:** Tasks 2–6 cover F1–F5; Task 7 covers all four Phase 1 items; Tasks 8–9 cover TP2–TP6; Tasks 10–11 cover every Phase 3 formula/data/analysis item; Tasks 12–13 cover all Phase 4 SDK and host-integration items; Tasks 14–15 enforce documentation, Roadmap state, interoperability, and final gates.
- **Dependency consistency:** the hard path and every parallel wave match the Mini-RFC dependency declarations. E3 precedes E1/E2; Host work follows Permission → Persistence → Comments → History → Collaboration → AI.
- **Type consistency:** `SpreadsheetDocument`, `DocumentController`, `FormulaProgram`, `CellPresentation`, `CompiledTemplate`, `GeneratedDocument`, `PrintDisplayList`, and adapter capability names remain consistent from producer to consumer tasks.
- **Scope control:** optional PDF/XLSX/Solver dependencies require recorded evaluation and remain tree-shakeable; host integrations remain protocols/UI connection points, not services.
- **Placeholder scan:** implementation steps name concrete behavior, files, commands, expected results, and acceptance gates; the plan contains no deferred design decisions.
