---
title: Workbook Data
---

# Workbook Data

`WorkbookInput` accepts either one sheet object or a complete `WorkbookData` array. Exported values are
always an ordered workbook array. Sheets use sparse JSON objects: row and cell keys are zero-based
decimal indexes, omitted entries are empty, and `len` records the logical extent.

```ts
import type { WorkbookData, WorkbookInput } from 'tego-sheet';

const input: WorkbookInput = [
  {
    name: 'Sparse example',
    rows: {
      len: 100,
      0: { cells: { 0: { text: 'A1' }, 3: { text: 'D1' } } },
      49: { cells: { 1: { text: 'B50' } } },
    },
    cols: { len: 12 },
  },
];

const serialized = JSON.stringify(input);
const restored = JSON.parse(serialized) as WorkbookData;
```

The key `"0"` identifies row 0 or column 0 even though spreadsheet labels display row 1 and column A.
For example, `rows["49"].cells["1"]` is cell B50. Coordinates in `CellPoint`, `CellAddress`, ranges,
selections, and events are also zero-based.

Workbook data can contain formulas, styles, merges, validations, filtering, frozen panes, row and
column dimensions, and JSON-compatible extension keys. Runtime `SheetId` values identify mounted
sheets for callbacks and commands; they are not serialized into workbook JSON.

Treat callback and `getValue()` results as immutable snapshots. Use `JSON.stringify` for persistence
and validate untrusted parsed data before passing it to a mounted component.

See the [WorkbookData API](/docs/api/type-aliases/WorkbookData) or load compatible sparse JSON in the
[Legacy JSON Playground](/playground?mode=legacy-json).
