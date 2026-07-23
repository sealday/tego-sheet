# Image adapter font decision

Date: 2026-07-23

Status: accepted for TP6 remediation

## Decision

Use `fontkit@2.0.4` as a direct runtime dependency of the optional image output
module. Load it dynamically only when a selected page uses a resolved
non-standard font.

The adapter uses Fontkit for shaping and glyph-outline extraction. It emits
only the outlines for glyphs used by the selected pages. It does not embed the
source font program, a full `@font-face`, CSS supplied by a font family, or an
external URL. The same converted path commands feed standalone SVG and the
Worker-safe canvas rasterizer, so SVG and PNG share geometry.

## Reproducible evidence

Run from the repository root:

```bash
npm view fontkit version license module browser exports dependencies dist.unpackedSize dist.integrity --json
```

Observed on 2026-07-23:

- version `2.0.4`;
- MIT license;
- browser and Node conditional exports;
- 5,610,637 unpacked bytes;
- integrity
  `sha512-syetQadaUEDNdxdugga9CpEYVaQIxOwk7GlwZWWZ19//qW4zE5bknOKeMBDYAASwnpaSHKJITRLMF9m1fp3s6g==`.

The upstream [Fontkit project](https://github.com/foliojs/fontkit) documents
browser operation, OpenType shaping, glyph layout, glyph paths, `path.toSVG()`,
and font subsetting. Its subset encoder is explicitly described as producing a
minimal PDF-oriented font without the standalone tables needed by browser
`@font-face`. Therefore the image adapter does not expose that subset as a web
font. Used-glyph outlines are the smaller and safer standalone SVG
representation.

## Security, license, and Worker boundary

- The adapter parses `OS/2.fsType` before shaping. Restricted, bitmap-only, and
  no-subsetting policies reject with a stable font error.
- Font family values are lookup keys only. They are never interpolated into
  CSS or markup.
- Missing glyphs reject instead of silently substituting.
- Fontkit receives only immutable bytes from `GeneratedDocument.resources`.
  It receives no URL, DOM node, controller, or editor state.
- Worker PNG consumes converted paths. It does not require `FontFace`,
  `document`, or an editor DOM. Missing `OffscreenCanvas`, `Path2D`, image
  decoding, or referenced resources rejects the whole request.
- Dynamic import keeps Fontkit out of the root entry and out of image requests
  which use only standard fonts.

## Upgrade trigger

Any version change must repeat glyph-shaping fixtures, CJK outline output,
browser/Worker PNG checks, package isolation, license-bit rejection, and
deterministic SVG checks.
