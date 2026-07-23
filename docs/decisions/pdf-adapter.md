# PDF adapter dependency decision

Date: 2026-07-23

Status: accepted for TP4

## Decision

Use `pdfkit@0.19.1` as a direct runtime dependency of the optional PDF output
module. Load it with a dynamic `import()` from the PDF adapter so consumers that
do not call PDF output do not execute the backend. Do not add `blob-stream`;
collect PDFKit's emitted `Uint8Array` chunks into one final `Blob` only after the
stream completes successfully.

The adapter remains the public boundary. PDFKit never receives a workbook,
controller, DOM, or layout callback. It receives only already-positioned
`GeneratedDocument.print.displayList` commands and already-resolved resources.

## Reproducible evidence

Run from the repository root:

```bash
npm view pdfkit version license main module dependencies dist.unpackedSize dist.integrity --json
npm view pdf-lib version license main module dependencies dist.unpackedSize dist.integrity --json
npm view @pdf-lib/fontkit version license main module dependencies dist.unpackedSize dist.integrity --json
```

Observed on 2026-07-23:

| Candidate        | Version  | License | Published package evidence                                                  | Result                |
| ---------------- | -------- | ------- | --------------------------------------------------------------------------- | --------------------- |
| PDFKit           | `0.19.1` | MIT     | ESM entry; 8,433,024 unpacked bytes; depends on current `fontkit ^2.0.4`    | selected              |
| pdf-lib          | `1.17.1` | MIT     | ESM entry; 19,461,112 unpacked bytes; last registry modification 2022-05-12 | rejected              |
| @pdf-lib/fontkit | `1.1.1`  | MIT     | ESM entry; 4,299,890 unpacked bytes; last registry modification 2022-04-06  | rejected with pdf-lib |

The upstream [PDFKit project](https://github.com/foliojs/pdfkit) states that it
works in Node and browsers and supports vector graphics, text, font embedding
and subsetting, PNG/JPEG images, annotations, and outlines. Its
[browser guide](https://pdfkit.org/docs/getting_started.html#using-pdfkit-in-the-browser)
documents browser `Blob` output. The dedicated documentation covers
[embedded font formats](https://pdfkit.org/docs/text.html#fonts),
[vector paths and clipping](https://pdfkit.org/docs/vector.html),
[link annotations](https://pdfkit.org/docs/annotations.html), and
[page outlines](https://pdfkit.org/docs/outline.html).

PDFKit's current font engine is Fontkit. The upstream
[Fontkit documentation](https://github.com/foliojs/fontkit) states that it runs
in Node and browsers, supports TrueType/OpenType/WOFF/WOFF2/TTC/dfont, performs
OpenType shaping, and creates PDF-oriented glyph subsets.

`pdf-lib` has a simpler byte-returning API and an ESM build, but custom fonts
require a separately registered `@pdf-lib/fontkit`. More importantly, its
maintainer explicitly records that
[subsetting does not work for all font types, including the reported Chinese case](https://github.com/Hopding/pdf-lib/issues/494).
That conflicts with TP4's required CJK-subset acceptance, so bundle convenience
cannot outweigh the correctness gap.

## Browser, Worker, bundle, and tree-shaking boundary

- PDFKit is JavaScript and its upstream supports browser execution without a
  DOM layout dependency. The adapter uses no `window` or `document`, so the same
  path is Worker-compatible.
- PDFKit is not assumed to be internally tree-shakeable. The adapter isolates it
  behind a dynamic import, producing an optional build chunk and keeping the
  ordinary spreadsheet, template, SVG, PNG, and XLSX paths free of PDFKit
  execution.
- The adapter does not use PDFKit's text wrapping or measurement. Every text
  command uses the fixed display-list coordinates with wrapping disabled.
- PDFKit's stream is buffered once. A `Blob` becomes observable only after
  `end`; abort, limit, font, image, or operator failures reject without a
  partial output.

## Font and determinism policy

- Standard PDF fonts are allowed only when every used code point is supported.
  CJK and other non-WinAnsi text requires a matching resolved font resource.
- The adapter inspects the OpenType `OS/2.fsType` embedding bits before giving a
  font to PDFKit. Restricted or bitmap-only embedding produces
  `PDF_FONT_EMBEDDING_FORBIDDEN`; absence or a subset failure produces
  `PDF_FONT_SUBSET_FAILED`. There is no silent fallback that could alter text
  geometry.
- PDFKit/fontkit performs glyph collection and subsetting. Acceptance parses the
  result and verifies searchable Unicode text plus a subset font descriptor for
  a redistributable CJK fixture.
- Metadata dates come only from `GeneratedDocument.metadata.generatedAt`;
  callers cannot introduce an implicit wall clock. Exact PDF bytes are not a
  public compatibility promise, but page order, geometry, metadata, selected
  pages, and text semantics are deterministic inputs.

## Security and limits

Only the display-list command allowlist is translated. Links permit `http`,
`https`, `mailto`, or internal fragment destinations. Metadata is length
bounded and stripped of control characters. Page count, resource bytes, output
bytes, and elapsed time are checked before returning. No network request,
controller read, HTML parsing, JavaScript action, attachment, encryption, form,
or arbitrary PDF operator is exposed.

## Upgrade trigger

Replace the backend without changing `PdfAdapter` if the pinned version fails a
supported CJK subset fixture, cannot run in the Worker package test, or adds a
non-MIT runtime dependency incompatible with this package. Any version change
must repeat the commands and acceptance matrix above.
