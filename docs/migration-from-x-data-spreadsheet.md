# Migration from x-data-spreadsheet

`tego-sheet` preserves supported workbook JSON and spreadsheet behavior, but intentionally provides a new React-only integration surface. Migrate application integration; do not rewrite stored workbook data.

## Removed legacy APIs

The old `new Spreadsheet(element, options)` constructor and `x_spreadsheet(...)` global do not exist. Render `<TegoSheet>` in React and import `tego-sheet/styles.css`. The old `.on(...)` emitter and string event names do not exist either; use typed props such as `onDocumentChange`, `onCellEdit`, `onPaste`, `onSelectionChange`, `onActiveSheetChange`, and `onError`.

```tsx
import { migrateLegacyWorkbook, TegoSheet, type SpreadsheetDocument } from 'tego-sheet';
import 'tego-sheet/styles.css';

interface EditorProps {
  legacyJson: unknown;
  saveDocument(document: SpreadsheetDocument): void;
}

export function Editor({ legacyJson, saveDocument }: EditorProps) {
  const migrated = migrateLegacyWorkbook(legacyJson);
  if (!migrated.ok) return <p>Legacy workbook could not be migrated.</p>;
  return <TegoSheet defaultDocument={migrated.document} onDocumentChange={saveDocument} />;
}
```

There is no public controller, renderer, `DataProxy`, mutable sheet object, global locale mutation,
internal emitter, string event API, or manual `destroy()` lifecycle. Imports from `src`, `core`,
`engine`, `controller`, or React internals are blocked by the package export map.

## Choose one ownership mode

- Uncontrolled: pass `defaultDocument`; use `onDocumentChange` for persistence and a `TegoSheetHandle` ref for queries or commands.
- Controlled: pass `document` and accept the next document from `onDocumentChange`. Keep the current object reference during unrelated parent renders.
- Never pass both props or switch modes after mount.

The input and callback payloads remain isolated from internal state. `getDocument()` returns an isolated schema 2 snapshot. Unknown JSON-compatible extension keys are preserved recursively through unrelated edits, history, and export.

## Map UI options to React props

| Legacy option or integration | React API                                                                                          |
| ---------------------------- | -------------------------------------------------------------------------------------------------- |
| `showToolbar: false`         | `toolbar={false}`                                                                                  |
| custom toolbar DOM           | `toolbar={renderer}` with typed `ToolbarRenderProps`                                               |
| `showBottomBar: false`       | `sheetTabs={false}`                                                                                |
| custom sheet-tab DOM         | `sheetTabs={renderer}` with typed `SheetTabsRenderProps`                                           |
| `mode: 'read'`               | `readOnly`                                                                                         |
| global locale registration   | `locale={definition}` on each component                                                            |
| imperative load              | controlled `document`, or remount an uncontrolled component with a new `key` and `defaultDocument` |
| imperative data export       | `ref.current.getDocument()` or the `onDocumentChange` document                                     |
| resize callbacks             | size the container and call `recalculateLayout()` when needed                                      |

Import bundled dictionaries from their explicit paths:

```tsx
import { de } from 'tego-sheet/locales/de';
import { en } from 'tego-sheet/locales/en';
import { nl } from 'tego-sheet/locales/nl';
import { zhCN } from 'tego-sheet/locales/zh-cn';
```

## Ref commands

Use `TegoSheetHandle` for `focus`, workbook and cell queries, cell edits, sheet management, undo/redo, validation, printing, and layout recalculation. Ref mutations follow the same validation, history, read-only, and callback rules as the default UI. Invalid caller commands throw the public `TegoSheetException` synchronously; recoverable browser failures are reported through `onError`.

## Five intentional correctness differences

These are the only approved visible departures from legacy behavior:

1. **Empty workbook:** loading `[]` renders an explicit empty workbook state and continues exporting `[]`; it no longer leaves a stale grid visible.
2. **All sheets are validated:** `validate()` inspects every sheet, including hidden rows and columns, in deterministic sheet/row/column order rather than validating only the initially created sheet.
3. **Rendered-value stable sort:** sorting uses the selected column's rendered values, handles numeric and textual groups deterministically, keeps empty values last in both directions, and preserves original row order for equal comparisons. It no longer sorts row indexes.
4. **Resource cleanup:** React unmount removes every owned listener, observer, timer, animation frame, subscription, and transient overlay. There is no incomplete manual destroy path.
5. **Printable cells:** cells with `printable: false` omit their content from print output while retaining grid geometry. The legacy implementation stored this flag but ignored it during printing.

## Data and identity notes

Sparse row, column, and cell indexes; formulas; styles; merges; validation; filters; falsy values; and extension keys retain their serialized meaning. Invalid workbook replacement is atomic, so partially imported state is never exposed.

## Migrate stored legacy workbooks to schema 2

Use `migrateLegacyWorkbook()` once at the persistence boundary. It accepts the former
single-sheet object or ordered sheet array and returns an atomic discriminated result:

```ts
import { migrateLegacyWorkbook, serializeSpreadsheetDocument } from 'tego-sheet';

const result = migrateLegacyWorkbook(legacyWorkbook);
if (!result.ok) {
  console.error(result.diagnostics);
  throw new Error('Legacy workbook migration failed');
}

const schema2Json = serializeSpreadsheetDocument(result.document);
```

For reproducible fixtures or migrations, provide deterministic `ids.documentId` and
`ids.sheetId` factories. Formulas remain source text and cached legacy values are never
treated as document truth. Features that schema 2 cannot represent are reported through
structured `LEGACY_FIELD_DEGRADED` or `LEGACY_FIELD_DROPPED` diagnostics rather than being
copied into an extension bag. Current row and column dimensions, hidden/default styles,
freeze panes, filters, validation, merges, and editable/printable cell flags are
persisted in normalized schema 2 fields.

Schema 2 persists stable sheet identity in `workbook.sheets[].id`. Callback and ref-command `SheetId`
values identify those sheets at runtime; serialize the document field rather than maintaining a
second identity store. External replacements retain the IDs they supply, while edits, renames,
history, and controlled acknowledgements preserve existing document IDs.

Each `MigrationDiagnostic` has a stable code, severity, message, and source location. Treat error
diagnostics as a failed import, persist warnings with the migration audit, and show informational
diagnostics only when requested. Migration is atomic: a failure never returns a partially
converted document.

## Compatibility boundary

| Input or behavior                                              | Migration result                                              |
| -------------------------------------------------------------- | ------------------------------------------------------------- |
| Ordered sheet arrays and single-sheet objects                  | Converted to schema 2 with stable sheet IDs                   |
| Sparse cells, formulas, styles, merges, filters and validation | Preserved in normalized fields                                |
| Cached formula values                                          | Dropped; the formula engine recalculates them                 |
| Unknown JSON-compatible extension keys                         | Preserved when valid schema extensions                        |
| Unsupported legacy fields                                      | Reported as `LEGACY_FIELD_DEGRADED` or `LEGACY_FIELD_DROPPED` |
| Functions, cyclic objects and unsafe prototype keys            | Rejected                                                      |

Enforce an application-level upload limit before parsing large imports. The library validates the
complete result before publication and never exposes partial streaming documents.

The package exposes only `tego-sheet`, `tego-sheet/styles.css`, the four locale subpaths, and `tego-sheet/package.json`. Imports from controller, engine, React internals, source, or legacy paths are unsupported and blocked by the export map.
