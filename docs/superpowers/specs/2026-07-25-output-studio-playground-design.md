# Output Studio Playground design

**Date:** 2026-07-25
**Status:** Approved for implementation planning

## Goal

Extend the documentation Playground with a complete, honest demonstration of tego-sheet's template
output pipeline:

1. edit a real spreadsheet template and print profile;
2. edit structured sample data;
3. compile and render one immutable `GeneratedDocument`;
4. preview the exact generated pages;
5. print or export PDF, PNG, and XLSX artifacts from that generated result.

The demonstration must make the component's distinguishing behavior visible. It is not a generic
download panel: users should understand that template inputs are compiled deterministically, that
diagnostics are first-class output, and that preview and output adapters do not inspect or re-layout
the editor DOM.

## Scope

### Included

- A top-level `Spreadsheet / Output Studio` workspace switch.
- The existing five spreadsheet presets without behavior changes.
- A prepared invoice document, template, print profile, sample data, and deterministic render
  environment.
- Restricted but real template editing through `TegoSheet` template mode and `TemplateDesigner`.
- Editable JSON sample data.
- Explicit apply-and-regenerate behavior.
- Exact generated-page preview.
- Browser printing.
- PDF, current-page PNG, and XLSX downloads.
- Visible compilation, pagination, output, and accessibility status.
- Desktop, intermediate, and narrow responsive layouts.

### Excluded

- A blank-template creation wizard.
- A general-purpose template IDE.
- Arbitrary file upload or user-provided remote resources.
- Background autosave or implicit debounce-based compilation.
- Multi-file PNG downloads, which browsers may block. The demo exports the selected page.
- New package dependencies.

## Information architecture

The Playground gains two top-level workspaces:

- **Spreadsheet** contains the existing ownership, custom chrome, locale, and migration presets.
- **Output Studio** contains the template generation and output demonstration.

This separation prevents editor configuration examples and document-generation workflows from
sharing one increasingly crowded preset row.

### URL contract

- `/playground?workspace=spreadsheet&mode=uncontrolled`
- `/playground?workspace=output`

Existing links containing only `?mode=<preset>` remain compatible and resolve to the Spreadsheet
workspace. Invalid workspace or mode values are canonicalized while preserving unrelated query
parameters. Browser back and forward navigation restores both workspace and spreadsheet mode.

## Output Studio layout

On wide screens, Output Studio uses three columns:

1. **Inputs**
   - invoice template and active print profile summary;
   - entry point to the restricted template workbench;
   - sample-data summary and JSON editor.
2. **Exact page preview**
   - generated revision and page metadata;
   - SVG pages rendered by the public `TemplatePreview`;
   - selected-page and zoom controls;
   - stale or blocked overlays when the draft and generated revision differ.
3. **Pipeline and outputs**
   - compile, bind, and paginate stages;
   - structured diagnostics;
   - Print, PDF, PNG, and XLSX actions;
   - per-action progress, cancellation, success, and failure state.

At intermediate widths the preview occupies the first row and input/output cards share the second
row. On narrow screens the order is preview, inputs, diagnostics, and outputs. Every interactive
control remains at least 44 CSS pixels tall.

## Component boundaries

### `PlaygroundShell`

Owns workspace URL state and renders the workspace navigation. It preserves the existing keyed reset
and error-boundary behavior without learning template or output details.

### `SpreadsheetWorkspace`

Contains the current mode picker, `PresetSession`, sheet panel, document JSON, and event inspector.
Moving this code is a behavior-preserving extraction covered by existing component and docs tests.

### `OutputStudio`

Composes the workbench, preview, pipeline status, diagnostics, and output actions. It does not
implement compilation or adapter behavior.

### `TemplateWorkbench`

Uses the public template surface:

- `TegoSheet mode="template"`;
- `template` and `onTemplateChange`;
- `TemplateDesigner`;
- explicit active print profile selection.

The prepared invoice is editable, but the workbench does not offer blank-project creation.

### `OutputPipeline`

Owns a small state machine:

- `ready`: committed inputs match the current generated document;
- `dirty`: draft template or data differs from the committed revision;
- `rendering`: the committed revision is compiling or rendering;
- `blocked`: the latest committed revision produced blocking diagnostics;
- `exporting`: represented per output action while the generated revision remains readable.

It stores draft inputs separately from committed inputs. `Apply & regenerate` validates JSON,
commits one new revision, aborts older work, compiles the template, and renders it. A late result may
commit only when its revision and abort signal still match the active request.

### `TemplatePreview`

Consumes the exact successful `GeneratedDocument`. It serializes the same page identities, sizes,
and display list later consumed by browser print, PDF, and image output. It performs no re-layout.

### `OutputActions`

Receives the current successful generated revision:

- `IsolatedBrowserPrintAdapter` prints its generated pages.
- `PdfAdapter` renders selected pre-paginated pages.
- `ImageAdapter` renders the currently selected page as PNG.
- `XlsxAdapter` writes the semantic workbook carried by the same `GeneratedDocument`.

The UI must not imply that XLSX consumes the page display list. It shares the generated document but
uses its semantic workbook representation.

### Fixtures and utilities

A dedicated fixture module owns the invoice source document, template, profile, sample data, and
render environment. Separate pure utilities own URL parsing, output filenames, diagnostic grouping,
and browser Blob download cleanup.

## Data and rendering flow

```text
draft template + draft JSON
            |
            | Apply & regenerate
            v
committed revision
            |
            v
compileSpreadsheetTemplate
            |
            v
renderSpreadsheetTemplate
            |
            +---- diagnostics
            |
            v
immutable GeneratedDocument
   |           |          |          |
preview      print       PDF/PNG     XLSX
display      display     display     semantic
list         list        list        workbook
```

The initial fixture is committed and rendered automatically on first entry, so the first view is a
working two-page example rather than an empty configuration screen.

## Editing and consistency rules

- Template and data edits modify drafts only.
- Any draft change marks the current preview stale and disables all output actions.
- `Apply & regenerate` is the only transition that commits draft inputs.
- There is no time-based debounce or hidden regeneration.
- A successful render atomically replaces the generated document and diagnostics.
- A blocked render preserves the prior successful preview with a stale/blocked overlay and disabled
  outputs. It never presents the prior artifact as the current revision.
- Reset restores the prepared invoice fixture and generates a new clean revision.

## Diagnostics and failures

- Diagnostics are grouped into blocking errors and non-blocking warnings.
- Blocking compile, binding, or pagination diagnostics identify their stable code and available
  source location.
- Invalid JSON remains in the editor, displays a parse error, and cannot commit.
- `PRINT_BLOCKED`, `RENDER_ABORTED`, output-limit, font, resource, and encoding failures are shown
  beside the initiating action.
- An output failure does not discard a valid generated document.
- Unexpected component render failures remain inside the existing Playground error boundary.
- Status changes are announced through an accessible live region.

## Output behavior

- Print explicitly states that it opens the system print dialog.
- Output buttons remain disabled for dirty, rendering, or blocked revisions.
- Each output action has independent busy and cancellation state.
- Deterministic names derive from the sample invoice identifier:
  - `invoice-INV-2026-042.pdf`
  - `invoice-INV-2026-042-page-1.png`
  - `invoice-INV-2026-042.xlsx`
- Blob URLs are revoked after the browser accepts the download action.
- Adapters and pending requests are disposed or aborted on unmount and revision replacement.

## Accessibility

- Workspace navigation uses a tablist or equivalent selected-state semantics.
- Template, sample data, preview, diagnostics, and outputs have named landmarks.
- Keyboard focus follows user actions but does not jump on background completion.
- Preview pages expose page labels and metadata in addition to SVG content.
- Stale, rendering, blocked, exporting, and completion states are announced.
- Controls meet the existing 44-pixel minimum target contract.
- Color is never the only diagnostic or state indicator.

## Testing strategy

Implementation follows red-green-refactor.

### Unit tests

- workspace and mode parsing/canonicalization;
- pipeline reducer transitions and revision ordering;
- stale-result and abort handling;
- diagnostic grouping;
- deterministic output filenames.

### Component tests

- Spreadsheet remains the compatible default workspace.
- Workspace changes update history and remount only the intended boundary.
- Output Studio renders the initial generated preview.
- Template and JSON edits mark the revision dirty and disable outputs.
- Apply commits exactly one revision and ignores late prior results.
- Invalid JSON and blocking diagnostics preserve the prior preview without enabling output.
- Template workbench uses the public template-mode props.
- Preview receives the same generated object later passed to output actions.
- Each adapter receives the current generated revision and reports success/failure accessibly.
- Blob URLs and adapters are cleaned up.

### Documentation browser tests

- Open Output Studio directly through its canonical URL.
- Exercise the initial happy path and one data edit/regeneration.
- Confirm the generated page count and output controls.
- Confirm keyboard access and live status.
- Validate the blocked diagnostic path without invoking a native print dialog.
- Cover desktop and narrow layouts.

### Visual tests

- Wide three-column Output Studio.
- Intermediate preview-first layout.
- Narrow single-column layout.
- Stale and blocked diagnostic states.

### Regression gates

- Existing Playground component, docs, and visual tests.
- TypeScript, lint, and formatting.
- Documentation build and static-server checks.
- Targeted output adapter tests; the Playground does not replace existing cross-browser or
  interoperability gates.

## Delivery constraints

- No new dependencies.
- No changes to public printing or output APIs are required.
- Existing Spreadsheet preset behavior remains stable.
- The demo must use public exports wherever a public surface exists.
- Automated browser tests must not open a real system print dialog.

## Success criteria

The feature is complete when a user can open Output Studio and understand, without reading source
code, that:

1. a real spreadsheet template and structured data generate a deterministic artifact;
2. diagnostics and page geometry are visible before output;
3. preview, print, PDF, and PNG share the same pre-paginated pages;
4. XLSX is generated from the semantic workbook in that same generated document;
5. changing inputs cannot accidentally export an older revision.
