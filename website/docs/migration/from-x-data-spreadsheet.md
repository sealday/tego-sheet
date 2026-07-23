---
title: From x-data-spreadsheet
---

# From x-data-spreadsheet

`tego-sheet` keeps supported sparse workbook JSON and spreadsheet behavior, but replaces the legacy
constructor, event emitter, and mutable internals with a React-only public API. Migrate integration
code; do not rewrite stored workbook data.

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
  return (
    <div style={{ height: 520 }}>
      <TegoSheet defaultDocument={migrated.document} onDocumentChange={saveDocument} />
    </div>
  );
}
```

| x-data-spreadsheet integration                              | tego-sheet React API                                                       |
| ----------------------------------------------------------- | -------------------------------------------------------------------------- |
| `new Spreadsheet(element, options)` or `x_spreadsheet(...)` | render `<TegoSheet>`                                                       |
| `.loadData(data)`                                           | migrate, then pass controlled `document` or remount with `defaultDocument` |
| `.getData()`                                                | `onDocumentChange` snapshots or `TegoSheetHandle.getDocument()`            |
| `.on('change', listener)`                                   | typed callback props such as `onDocumentChange` and `onCellEdit`           |
| `showToolbar: false`                                        | `toolbar={false}`                                                          |
| `showBottomBar: false`                                      | `sheetTabs={false}`                                                        |
| `mode: 'read'`                                              | `readOnly`                                                                 |
| global locale registration                                  | per-instance `locale` with an explicit locale subpath import               |
| custom toolbar or tabs DOM                                  | typed `ToolbarRenderProps` and `SheetTabsRenderProps` renderers            |
| resize callback                                             | a sized parent and `TegoSheetHandle.recalculateLayout()` when needed       |

There is no public controller, renderer, mutable sheet object, `DataProxy`, internal emitter, or manual
destroy method. Choose controlled or uncontrolled ownership at mount and do not switch modes while the
instance remains mounted. Schema 2 persists stable identity in `workbook.sheets[].id`; callback and
ref-command `SheetId` values identify those document sheets at runtime.

For the exact compatibility surface, preserved extension keys, ref command rules, and the five
intentional correctness differences, read the package's canonical
[migration document on GitHub](https://github.com/sealday/tego-sheet/blob/main/docs/migration-from-x-data-spreadsheet.md).

Next, review [Workbook Data](../concepts/workbook-data), the
[TegoSheet API](/docs/api/variables/TegoSheet), or the
[Legacy JSON Playground](/playground?mode=legacy-json).
