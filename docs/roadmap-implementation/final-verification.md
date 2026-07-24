# Product Roadmap final verification

Verification date: 2026-07-24

This ledger records observed evidence. A capability remains planned when its Mini-RFC requires an
external or manual gate that has not run, even when its implementation and local tests pass.

## Automated gates

| Gate                                            | Result              | Evidence                                                                        |
| ----------------------------------------------- | ------------------- | ------------------------------------------------------------------------------- |
| Foundation and template-printing focused suites | pass                | Recorded in shipped roadmap evidence                                            |
| XLSX DrawingML focused suites                   | pass                | Independent review approved; 28/28 final focused tests                          |
| Analysis focused suites                         | pass                | Final Table/visualization production probes 53/53; legacy projection 6/6        |
| Host integration contract suites                | pass                | Final contract and React integration matrix 88/88                               |
| SDK focused suites                              | pass                | 58/58 across cells, templates and adapters                                      |
| Package consumer gate                           | pending final rerun | Earlier 47/47; rerun after the final roadmap and architecture reconciliation    |
| TypeScript, lint and format                     | pass                | Final integrated `format:check`, `lint`, and `typecheck` completed successfully |
| Browser, visual, SSR, demo and docs             | pending final rerun | Run from the final clean worktree                                               |

## Interoperability matrix

| Surface                                    | Result              | Notes                                                                                              |
| ------------------------------------------ | ------------------- | -------------------------------------------------------------------------------------------------- |
| OpenPyXL structural probe                  | pass                | Formula, literal string, print area, paper size, conditional formats and sheet state               |
| LibreOffice                                | pass                | Headless XLSX round-trip; `veryHidden` becomes `hidden`, inch paper sizes normalize to millimetres |
| Excel Desktop                              | blocked             | Installed application was detected, but the opt-in AppleScript probe timed out                     |
| Excel for web                              | not run             | No authorized upload/account session is available                                                  |
| Browser print: Chrome/Firefox/Safari       | pending final rerun | Required release matrix                                                                            |
| PDF/image browser and Worker               | pending final rerun | Required release matrix                                                                            |
| Accessibility: keyboard and screen readers | pending manual gate | VoiceOver plus one additional screen reader/browser pair required                                  |

The XLSX implementation gate passes, but the `xlsx-output` and `file-interchange` roadmap items
cannot be marked shipped until Excel Desktop and Excel for web evidence is recorded.

## Independent review

| Area                            | Verdict                    | Remaining condition                                     |
| ------------------------------- | -------------------------- | ------------------------------------------------------- |
| Screen object rendering         | APPROVE                    | None                                                    |
| XLSX DrawingML                  | APPROVE                    | External Office interoperability remains a release gate |
| Analysis, Tables and visualizer | pending final confirmation | Parallel Table row projection received one final fix    |
| Host integrations and security  | APPROVE                    | None                                                    |

## Final reconciliation rules

- Do not infer manual Office, browser, visual, accessibility, license, or security evidence from unit tests.
- A local adapter implementation may be complete while its public roadmap item remains planned.
- Every shipped state must match `website/src/data/roadmap.ts`, the acceptance ledger, the roadmap
  index, and the shipped record.
- The final clean-worktree gate must run after all implementation and documentation commits.
