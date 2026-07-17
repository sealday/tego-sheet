---
title: From x-data-spreadsheet
---

# From x-data-spreadsheet

`tego-sheet` keeps supported sparse workbook JSON and spreadsheet behavior, but replaces the legacy
constructor, event emitter, and mutable internals with a React-only public API. Migrate integration
code; do not rewrite stored workbook data.

```tsx
import { TegoSheet, type WorkbookData, type WorkbookInput } from 'tego-sheet';
import 'tego-sheet/styles.css';

interface EditorProps {
  legacyJson: WorkbookInput;
  saveWorkbook(value: WorkbookData): void;
}

export function Editor({ legacyJson, saveWorkbook }: EditorProps) {
  return (
    <div style={{ height: 520 }}>
      <TegoSheet defaultValue={legacyJson} onChange={saveWorkbook} />
    </div>
  );
}
```

| x-data-spreadsheet integration                              | tego-sheet React API                                                 |
| ----------------------------------------------------------- | -------------------------------------------------------------------- |
| `new Spreadsheet(element, options)` or `x_spreadsheet(...)` | render `<TegoSheet>`                                                 |
| `.loadData(data)`                                           | controlled `value`, or remount with a new `key` and `defaultValue`   |
| `.getData()`                                                | `onChange` snapshots or `TegoSheetHandle.getValue()`                 |
| `.on('change', listener)`                                   | typed callback props such as `onChange` and `onCellEdit`             |
| `showToolbar: false`                                        | `toolbar={false}`                                                    |
| `showBottomBar: false`                                      | `sheetTabs={false}`                                                  |
| `mode: 'read'`                                              | `readOnly`                                                           |
| global locale registration                                  | per-instance `locale` with an explicit locale subpath import         |
| custom toolbar or tabs DOM                                  | typed `ToolbarRenderProps` and `SheetTabsRenderProps` renderers      |
| resize callback                                             | a sized parent and `TegoSheetHandle.recalculateLayout()` when needed |

There is no public controller, renderer, mutable sheet object, `DataProxy`, internal emitter, or manual
destroy method. Choose controlled or uncontrolled ownership at mount and do not switch modes while the
instance remains mounted. `SheetId` is runtime-only and must not be persisted in workbook JSON.

For the exact compatibility surface, preserved extension keys, ref command rules, and the five
intentional correctness differences, read the package's canonical
[migration document on GitHub](https://github.com/sealday/tego-sheet/blob/main/docs/migration-from-x-data-spreadsheet.md).

Next, review [Workbook Data](../concepts/workbook-data), the
[TegoSheet API](/docs/api/variables/TegoSheet), or the
[Legacy JSON Playground](/playground?mode=legacy-json).
