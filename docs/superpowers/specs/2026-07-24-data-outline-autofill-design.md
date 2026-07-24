# DATA-01 Outline Grouping and Enhanced Autofill Design

## Scope

This slice adds persistent row/column outline groups and revision-bound enhanced autofill. It
does not add editor chrome or change template-owned files.

## Outline model

Each worksheet owns a canonical `groups` collection. A group contains a stable opaque `id`, an
`axis` (`row` or `column`), inclusive `start` and `end` indexes, a normalized one-based `level`,
and `collapsed`. IDs are unique within a worksheet. Ranges are non-empty, bounded by the
worksheet logical size when supplied, and may be disjoint or properly nested. Crossing ranges
on the same axis are invalid. Canonical order is axis, start, end descending, then ID.

`group`, `ungroup`, and `toggle-group` are versioned public commands. They pass through the
existing validate → authorize → schema-plan → transaction/history pipeline. Group creation
derives levels from containment rather than trusting callers. Ungroup removes one stable ID;
toggle changes only `collapsed`.

Structural insert/delete transforms group endpoints with the shared coordinate transform.
Empty groups are deleted, duplicate/invalid groups are rejected, and levels are recomputed
after every structural change.

Collapsed groups do not overwrite `SheetRow.hidden` or `SheetColumn.hidden`. Runtime, display,
print, and export projections use the union of explicit hidden entries and indexes covered by
collapsed groups, so expanding a group restores the original row/column truth.

## Enhanced autofill

Autofill planning is revision-bound and immutable. The planner inspects a bounded source and
target, yields every 256 cells, and publishes one normalized `autofill` command. Commit uses one
document transaction.

Schema autofill owns typed-input semantics. Literal typed inputs are copied without display-text
round-tripping. Formula inputs are parsed into the public formula AST, translated by row/column
delta, and rendered back to canonical formula source. Scalar numeric, ISO-date, and numeric
suffix seeds infer a deterministic arithmetic pattern along the fill axis; other inputs tile.
The target command count/cell count and preview samples are bounded before publication.

## Stable failures

Outline resource or nesting violations fail with `GROUP_LIMIT_EXCEEDED`. Invalid command shapes
remain `COMMAND_SCHEMA_INVALID`. Autofill uses the existing `TRANSFORM_TOO_LARGE`,
`TRANSFORM_ABORTED`, `TRANSFORM_PLAN_STALE`, and `TRANSFORM_PLAN_MISSING` failures.

## Verification

Tests cover parsing/canonicalization, command permissions and history, structural transforms,
explicit-hidden union semantics in runtime and print projections, typed and AST formula
autofill, pattern inference, stale plans, budgets, and post-invocation cancellation.
