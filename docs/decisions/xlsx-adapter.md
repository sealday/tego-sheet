# XLSX adapter dependency decision

Date: 2026-07-23

Status: accepted for TP5

## Decision

Implement a project-owned, write-only OOXML serializer for the explicitly
supported Workbook 2.0 surface and use only `fflate@0.8.3` for deterministic ZIP
packaging. Import the named `zipSync` and `strToU8` exports. Fix every entry's
mtime to the ZIP epoch and insert entries in canonical path order.

Do not use ExcelJS or SheetJS Community Edition. Their broad mutable workbook
models and read/round-trip surfaces are outside this adapter's needs, while TP5
requires exact allowlisted XML, deterministic relationship targets and package
ordering, and stable diagnostics for every unsupported source feature.

## Reproducible evidence

Run from the repository root:

```bash
npm view fflate version license main module browser exports sideEffects dependencies dist.unpackedSize dist.integrity --json
npm view exceljs version license main browser dependencies dist.unpackedSize dist.integrity --json
npm view xlsx version license main module browser sideEffects dependencies dist.unpackedSize dist.integrity --json
```

Observed on 2026-07-23:

| Candidate                       | Version  | License    | Published package evidence                                                                                     | Result                |
| ------------------------------- | -------- | ---------- | -------------------------------------------------------------------------------------------------------------- | --------------------- |
| fflate                          | `0.8.3`  | MIT        | browser/Node conditional exports, ESM, `sideEffects: false`, no runtime dependencies, 796,742 unpacked bytes   | selected for ZIP only |
| ExcelJS                         | `4.4.0`  | MIT        | browser bundle; 21,825,509 unpacked bytes; dependencies include JSZip, archiver, streams, CSV and unzip stacks | rejected              |
| SheetJS `xlsx` registry package | `0.18.5` | Apache-2.0 | ESM and `sideEffects: false`; 7,499,035 unpacked bytes plus legacy multi-format dependencies                   | rejected              |

The upstream [fflate project](https://github.com/101arrowz/fflate) documents
pure-JavaScript browser and Node operation, ESM tree shaking, ZIP support,
Unicode filenames, Worker-backed asynchronous APIs, and explicit `mtime`
control. It reports about 7 kB minified for ZIP compression when tree shaken.
The adapter uses the synchronous named exports because output is already
expected to run inside a host-selected main thread or Worker and synchronous
ordering makes cancellation checkpoints and deterministic packaging explicit.

The upstream [ExcelJS project](https://github.com/exceljs/exceljs) supports
writing XLSX values, styles, merges, data validation, conditional formatting,
images, and page setup. It is feature-capable, but its published browser entry
is a monolithic minified bundle and its dependency graph includes both Node
archive/stream and browser ZIP paths. It does not make deterministic byte
ordering and timestamps part of its public writer contract.

[SheetJS Community Edition](https://docs.sheetjs.com/) focuses on broad
spreadsheet format extraction and generation. Its public CE support surface
does not provide the narrow, fully controlled style, validation, conditional
formatting, drawing, and print-part contract TP5 requires. The registry package
also brings parsers for unrelated legacy formats. A write-only allowlist is
smaller and safer than exposing those read and round-trip capabilities.

## OOXML contract

The serializer emits only these project-owned part families when needed:

- `[Content_Types].xml`, root relationships, workbook, workbook relationships,
  workbook properties, shared strings, styles, and ordered worksheets;
- worksheet cells with explicit types, formulas and cached values, merges,
  row/column dimensions and hidden state, data validation, supported
  conditional-format rules, print area/titles, page margins/setup, row/column
  breaks, and header/footer;
- drawing, drawing relationships, and validated PNG/JPEG media parts.

All element names, attributes, relationship types, and targets are emitted by
code constants. User strings pass through XML escaping; relationship targets
are generated identifiers, never user-provided paths. No arbitrary XML,
external relationship, macro, OLE, script, connection, or unknown imported
part can enter the package.

## Determinism, browser, Worker, and tree-shaking

- Worksheets, cells, merges, styles, validations, strings, resources,
  relationships, and ZIP paths use stable source or canonical key order.
- ZIP uses fixed `1980-01-01T00:00:00.000Z` mtimes, one compression level, and
  canonical path insertion. Equal document plus equal options must produce
  equal bytes, not merely equivalent XML.
- `fflate` has browser and Node conditional exports and no DOM dependency. The
  serializer itself uses `Uint8Array`, `TextEncoder`, and `Blob`, all available
  in the supported Node/browser/Worker environments.
- Named ESM imports and `sideEffects: false` keep decompression, streaming, GZIP,
  and unrelated codecs out of the output chunk.

## Unsupported-feature policy

Workbook 2.0 currently has no first-class conditional-formatting or full
header/footer schema. The adapter never invents those semantics. Recognized
future metadata is serialized only after its schema is promoted into the
document contract; otherwise a located `XLSX_UNSUPPORTED_FEATURE` diagnostic
blocks export when omission would violate a declared feature.

Custom cells, unsupported validation/style JSON, unsupported image MIME,
invalid sheet names, formulas that cannot map to Excel syntax, and limit
violations produce their dedicated stable diagnostic. Plain strings beginning
with `=` remain strings. No unsupported source feature is silently converted to
OOXML.

## Upgrade trigger

Any change to `fflate` or adoption of a higher-level workbook library requires
re-running the metadata commands, deterministic-byte fixtures, browser/Worker
package tests, and Excel/LibreOffice structure checks. The public adapter and
diagnostic contract does not depend on the ZIP implementation.
