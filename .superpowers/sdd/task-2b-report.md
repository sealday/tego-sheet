# Task 2B Report: Legacy workbook migration

## Outcome

Implemented the pure one-shot `migrateLegacyWorkbook()` boundary and exported its
runtime and public result/options/diagnostic types from `src/document` and the package
root. The migrator does not import the controller, React, DOM, formula runtime, network
code, or a runtime dependency. Its candidate output is accepted only after
`parseSpreadsheetDocument()` validates and freezes it.

## RED evidence

After adding the tracked fixtures and `tests/unit/document/legacy-migration.test.ts`:

```text
npx vitest run --project unit tests/unit/document/legacy-migration.test.ts
Test Files  1 failed (1)
Tests       9 failed (9)
TypeError: migrateLegacyWorkbook is not a function
```

The failure was the intended missing-runtime failure, before production implementation.

## GREEN and gate evidence

Final brief gate run:

```text
npx vitest run --project unit tests/unit/document
Test Files  3 passed (3)
Tests       59 passed (59)

npx vitest run --project unit tests/unit/bootstrap.test.ts \
  --project architecture tests/architecture/public-surface.test.ts \
  tests/architecture/public-api-documentation.test.ts \
  tests/architecture/core-purity.test.ts
Test Files  4 passed (4)
Tests       14 passed (14)

npm run test:package
tests 40; pass 40; fail 0

npm run typecheck
exit 0

npm run lint
exit 0

npm run format:check
All matched files use the correct format.
```

The first package run correctly exposed that its tracked-source fixture omits untracked
files; staging the new migrator made that fixture representative. A subsequent
architecture run found a forbidden legacy type name in JSDoc; the wording was removed.
The complete gate sequence then passed.

## Fixture and behavior coverage

- Single-sheet and ordered multi-sheet legacy inputs.
- Stable injected document and sheet ID factories and deterministic serialized bytes.
- Sparse numeric coordinates, including row 10,000 and column 52.
- Formula source without evaluation or cached-value promotion.
- Explicit `''`, `0`, `false`, and explicit blank cells.
- Canonical style deduplication, nested unknown-style-field reporting, and style refs.
- Normalized A1 merges and reliable validation refs, including explicit blank targets.
- Dimensions, hidden state, row/column styles, freeze, filters, editable, and printable
  diagnostics when schema 2 has no corresponding persistent field.
- Unknown fields are reported and omitted; `extensions` remains empty.
- Invalid values/references and JavaScript `Date` values fail atomically.
- Input mutation isolation, output freezing, and no shared style references.
- Root runtime exports, public TypeDoc graph, bootstrap, package consumer, and migration
  guide contract.

## Files changed

- `src/document/migrate-legacy.ts`
- `src/document/index.ts`
- `src/index.ts`
- `tests/unit/document/legacy-migration.test.ts`
- `tests/fixtures/document/legacy/complete.json`
- `tests/fixtures/document/legacy/multiple-sheets.json`
- `tests/unit/bootstrap.test.ts`
- `tests/architecture/public-surface.test.ts`
- `tests/architecture/public-api-documentation.test.ts`
- `tests/package/package-exports.test.mjs`
- `docs/migration-from-x-data-spreadsheet.md`

## Decisions and self-review

- Legacy style and validation IDs are derived deterministically from canonical first
  occurrence; only document/sheet IDs require injected factories for deterministic bytes.
- Validation refs are expanded to explicit schema 2 cells, capped at one million cells
  per ref before allocation.
- Cached legacy `value` is ignored whenever text supplies document truth.
- No locale-sensitive comparator, formula evaluator, Date serialization, extension bag,
  or caller-owned object is used in the returned document.
- Failures expose diagnostics only; the parser remains the final schema authority.

## Commit

`feat(document): migrate legacy workbooks` (SHA is reported by the implementing agent
after the commit is created).

## Concerns

Task 2A's schema 2 `Sheet` currently has no persistent fields for row/column dimensions,
hidden/default styles, freeze, autofilter, or legacy editable/printable state. Task 2B
therefore reports each such known field as `LEGACY_FIELD_DEGRADED` and deliberately does
not hide it in `extensions`. A future schema task must add normalized fields before these
features can be converted without loss.

## Review-fix addendum

The independent Task 2B review returned `REQUEST_CHANGES`. The original concern above
is resolved by the follow-up commit rather than deferred: schema 2 now persists sparse
row/column layout and default style references, logical counts, freeze coordinates,
normalized filter/sort state, and cell editable/printable flags. Parser validation,
canonical ordering, row/column limits, public types, TypeDoc, serializer round trips,
package declarations, and migration docs cover these additions.

### Review RED evidence

Focused schema and migration regressions initially produced 13 failures:

- layout/filter/print fields were absent after parse/serialize;
- authoritative text with a cached `Date` failed;
- duplicate sheet/cell merge representations reached `INVALID_MERGE`;
- leading-zero keys and sparse collection extensions failed;
- invalid style, validation, and cell discriminator semantics succeeded;
- cumulative validation expansion reached the parser only after allocation;
- an accessor-backed legacy object threw past `MigrationResult`.

A separate validation-limit RED proved `maxRows`/`maxColumns` were not enforced, and a
validation-payload RED proved an invalid list payload was accepted.

### Review GREEN evidence

The final fresh gate run after all review fixes:

```text
npx vitest run --project unit tests/unit/document
Test Files  3 passed (3)
Tests       86 passed (86)

npx vitest run --project unit tests/unit/bootstrap.test.ts \
  --project architecture tests/architecture/public-surface.test.ts \
  tests/architecture/public-api-documentation.test.ts \
  tests/architecture/core-purity.test.ts
Test Files  4 passed (4)
Tests       14 passed (14)

npm run test:package
tests 40; pass 40; fail 0

npm run typecheck
npm run lint
npm run format:check
all exit 0
```

The fixture matrix now directly migrates all 11 current workbook serialization inputs:
autofilter, blank object, cells, columns, empty array, multiple sheets, rows, sheet
fields, sparse falsy/extensions, styles, and validations.

### Review decisions

- Exact duplicate merge ranges from sheet and anchor-cell representations are
  deduplicated; distinct overlaps remain for the schema parser to reject.
- Legacy input is captured once from own data-property descriptors. Accessors,
  descriptor failures, and cycles return located atomic migration failures.
- Leading-zero decimal keys normalize before collision detection. Nonnumeric collection
  extensions produce located drop warnings and are never copied.
- Known legacy style, validation type/operator/payload, cell discriminator, layout,
  filter, editable, and printable values receive semantic validation before generic
  schema registries can accept them.
- Validation expansion uses a cumulative one-million-cell work limit before allocation;
  the schema parser remains the final authority afterward.
- Cached values, including `Date`, are ignored when authoritative `text` exists.

### Follow-up commit

`fix(document): preserve legacy workbook semantics` (SHA reported after commit creation).

### Remaining concerns

None known within Task 2B scope.

## Re-review edge-semantics addendum

The Task 2B re-review identified three boundary contracts absent from the prior fixture
matrix. New RED regressions produced eight focused failures before implementation:

- a valid comma-separated list validation was rejected;
- an explicit empty-string filter value failed both migration and direct schema round trip;
- negative font size, row height, and column width passed migration, while the two
  normalized layout dimensions also passed direct schema parsing.

The follow-up normalizes a legacy list string with the current runtime's comma-splitting
semantics into a schema 2 array payload. Filter values use the possibly-empty display
string decoder. Font and layout geometry must be finite and non-negative in migration,
and the schema parser independently enforces the normalized row/column invariant.

`tests/fixtures/document/legacy/edge-semantics.json` tracks the supported boundaries:
comma-separated list data, an empty filter item, and zero-valued font/row/column
geometry. Focused GREEN after the changes was 94/94 document tests.

The edge-semantics fix is committed as
`fix(document): align legacy edge semantics` (SHA reported after commit creation).
