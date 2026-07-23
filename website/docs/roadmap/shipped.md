# Shipped Roadmap capabilities

Roadmap capabilities move here only after their acceptance-ledger implementation tasks and
verification commands pass.

| Phase | Capability                                               | Status  | Design                                    |
| ----- | -------------------------------------------------------- | ------- | ----------------------------------------- |
| 0     | Workbook 2.0 typed document model                        | shipped | [Foundation](foundation.md)               |
| 0     | Atomic Command / Transaction                             | shipped | [Foundation](foundation.md)               |
| 0     | Formula dependency and number-format core                | shipped | [Foundation](foundation.md)               |
| 0     | Shared render semantics and Canvas accessibility         | shipped | [Foundation](foundation.md)               |
| 0     | Minimal cell-type and adapter registry kernel            | shipped | [Foundation](foundation.md)               |
| 1     | Sheet, selection and range print targets                 | shipped | [Template printing](template-printing.md) |
| 1     | Safe scalar bindings, repeat rows and conditional ranges | shipped | [Template printing](template-printing.md) |
| 1     | Deterministic pagination and print profiles              | shipped | [Template printing](template-printing.md) |
| 1     | Page preview and isolated browser printing               | shipped | [Template printing](template-printing.md) |

Workbook 2.0 completed its schema boundary, legacy migration, runtime-controller integration,
React ingress migration, and command-based rich-cell transforms on 2026-07-23. Its acceptance
ledger and focused, architecture, package, SSR, and browser suites passed before shipment.

Atomic Command / Transaction shipped on 2026-07-23 with versioned public commands, atomic
multi-command commits, revision and permission gates, dry-run previews, minimal forward/inverse
patch history, unified coordinate transforms, multi-sheet change summaries, and observer-safe
commit notifications. Generated transaction invariants, public API, package, SSR, and docs gates
passed before shipment.

Formula dependency and number-format core shipped on 2026-07-23 with typed formula values and ASTs,
stable references, bounded incremental dependency calculation, deterministic runtime environments,
explicit function compatibility, Excel-oriented number/date/time formats, and typed Canvas/sort
projection. Formula, legacy compatibility, architecture, package, SSR, and build gates passed
before shipment.

Shared render semantics and Canvas accessibility shipped on 2026-07-23 with one bounded,
cache-budgeted `CellPresentation` path for Canvas, semantic DOM, validation, and display-list
printing. Production integration, focus/edit semantics, font metrics, package, SSR, and
cross-browser print parity gates passed before shipment.

The minimal extension kernel shipped on 2026-07-23 with declaration-mergeable capability typing,
host-environment-bound deterministic resolution, safe manifest snapshots, and reentrancy-safe
initialize/unregister/dispose lifecycle handling. Its adversarial lifecycle and cell-semantics
suites passed before shipment.

The Template Print MVP shipped on 2026-07-23 with safe scalar, repeat-row, and conditional
bindings; ordered sheet and range targets; deterministic pagination and print profiles; shared
SVG preview and isolated browser printing; and template-mode Canvas decorations. Compiler,
resource-budget, component, package, SSR, docs, and cross-browser print gates passed before
shipment.
