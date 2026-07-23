# Shipped Roadmap capabilities

Roadmap capabilities move here only after their acceptance-ledger implementation tasks and
verification commands pass.

| Phase | Capability                                    | Status  | Design                      |
| ----- | --------------------------------------------- | ------- | --------------------------- |
| 0     | Workbook 2.0 typed document model             | shipped | [Foundation](foundation.md) |
| 0     | Minimal cell-type and adapter registry kernel | shipped | [Foundation](foundation.md) |

Workbook 2.0 completed its schema boundary, legacy migration, runtime-controller integration,
React ingress migration, and command-based rich-cell transforms on 2026-07-23. Its acceptance
ledger and focused, architecture, package, SSR, and browser suites passed before shipment.

The minimal extension kernel shipped on 2026-07-23 with declaration-mergeable capability typing,
host-environment-bound deterministic resolution, safe manifest snapshots, and reentrancy-safe
initialize/unregister/dispose lifecycle handling. Its adversarial lifecycle and cell-semantics
suites passed before shipment.
