# Shipped Roadmap capabilities

Roadmap capabilities move here only after their acceptance-ledger implementation tasks and
verification commands pass.

| Phase | Capability                                                   | Status  | Design                                                  |
| ----- | ------------------------------------------------------------ | ------- | ------------------------------------------------------- |
| 0     | Workbook 2.0 typed document model                            | shipped | [Foundation](foundation.md)                             |
| 0     | Atomic Command / Transaction                                 | shipped | [Foundation](foundation.md)                             |
| 0     | Formula dependency and number-format core                    | shipped | [Foundation](foundation.md)                             |
| 0     | Shared render semantics and Canvas accessibility             | shipped | [Foundation](foundation.md)                             |
| 0     | Minimal cell-type and adapter registry kernel                | shipped | [Foundation](foundation.md)                             |
| 1     | Sheet, selection and range print targets                     | shipped | [Template printing](template-printing.md)               |
| 1     | Safe scalar bindings, repeat rows and conditional ranges     | shipped | [Template printing](template-printing.md)               |
| 1     | Deterministic pagination and print profiles                  | shipped | [Template printing](template-printing.md)               |
| 1     | Page preview and isolated browser printing                   | shipped | [Template printing](template-printing.md)               |
| 2     | Nested, horizontal, range and page repeats                   | shipped | [Template printing](template-printing.md)               |
| 2     | Image, font, QR code and async resource pipeline             | shipped | [Template printing](template-printing.md)               |
| 2     | PDF Blob output                                              | shipped | [Template printing](template-printing.md)               |
| 2     | SVG and PNG page output                                      | shipped | [Template printing](template-printing.md)               |
| 3     | Conditional formatting                                       | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Advanced validation, dropdown and checkbox cells             | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Expanded function library and cross-sheet references         | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Named ranges, array and spill formulas                       | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Multi-column sort, conditional filter and saved views        | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Grouping, deduplication, text split and data cleanup         | shipped | [Formulas and data](formulas-data.md)                   |
| 3     | Structured tables and structured references                  | shipped | [Analysis and visualization](analysis-visualization.md) |
| 3     | Charts and Sparklines                                        | shipped | [Analysis and visualization](analysis-visualization.md) |
| 3     | Images, shapes, text boxes and anchored objects              | shipped | [Analysis and visualization](analysis-visualization.md) |
| 3     | PivotTable and Slicer                                        | shipped | [Analysis and visualization](analysis-visualization.md) |
| 3     | Goal Seek and pluggable Solver                               | shipped | [Analysis and visualization](analysis-visualization.md) |
| 4     | Public structured cell renderer/editor plugin SDK            | shipped | [Extensibility](extensibility.md)                       |
| 4     | Versioned Template Module SDK                                | shipped | [Extensibility](extensibility.md)                       |
| 4     | Public adapter lifecycle, trust policy and compatibility SDK | shipped | [Extensibility](extensibility.md)                       |
| 4     | Persistence and version history adapters                     | shipped | [Host integrations](host-integrations.md)               |
| 4     | Collaboration and remote selection adapters                  | shipped | [Host integrations](host-integrations.md)               |
| 4     | Permission and comment adapters                              | shipped | [Host integrations](host-integrations.md)               |
| 4     | Validated AI command proposals                               | shipped | [Host integrations](host-integrations.md)               |

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

Advanced template structures and the resource pipeline shipped on 2026-07-23 with nested row,
column, two-dimensional, page, sheet, and subtemplate expansion; stable structural and object
mappings; clone-before-allocation limits; bounded image, font, SVG, and standard QR resolution;
session-owned cleanup; and resolved-resource display-list output. Their adversarial axis,
resource-safety, package, TypeDoc, and print regression gates passed before shipment.

PDF and image output shipped on 2026-07-23 as isolated optional adapters that translate the
immutable generated display list without re-layout. PDF includes searchable vector text, licensed
CJK glyph subsets, deterministic metadata, page selection, links, and bounded generation. Image
output provides safe standalone SVG and DPI-accurate PNG in browsers and Workers. Their package,
browser, Worker, visual, font, cancellation, and resource-isolation gates passed before shipment.

The XLSX adapter implementation is complete and locally verified, but its Roadmap item remains
planned until the recorded Excel Desktop and Excel Web external interoperability gates pass.

Spreadsheet depth and the public SDK ecosystem shipped on 2026-07-24. The release includes
derived conditional presentation, advanced validation and structured cells, the expanded formula
engine, document and session views, atomic data cleanup, stable structured tables, persisted
charts and sparklines, anchored objects, pivot/slicer analysis, bounded Goal Seek and Solver
contracts, versioned cell/template/adapter SDKs, and host-owned persistence, history,
collaboration, permission, comments, and validated AI proposal workflows. Focused hostile-input,
screen/print parity, package, architecture, security, and public API gates passed before shipment.

Common file interchange remains planned with XLSX output because the required Excel Desktop and
Excel for web external interoperability evidence is not yet available.
