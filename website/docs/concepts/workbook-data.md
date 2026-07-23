---
title: Workbook Data
---

# Workbook Data

`SpreadsheetDocument` is the only runtime and persistence model. It carries an explicit
`schemaVersion`, stable document and sheet IDs, registries, resources, and an ordered sheet list.

```ts
import { createSpreadsheetDocument, parseSpreadsheetDocument } from 'tego-sheet';

const document = createSpreadsheetDocument({ sheetName: 'Sparse example' });

const serialized = JSON.stringify(document);
const restored = parseSpreadsheetDocument(JSON.parse(serialized));
if (!restored.ok) throw new Error('Stored document is invalid');
```

The key `"0"` identifies row 0 or column 0 even though spreadsheet labels display row 1 and column A.
For example, `rows["49"].cells["1"]` is cell B50. Coordinates in `CellPoint`, `CellAddress`, ranges,
selections, and events are also zero-based.

Workbook data can contain formulas, styles, merges, validations, filtering, frozen panes, row and
column dimensions, and JSON-compatible extension keys. Runtime `SheetId` values identify mounted
sheets for callbacks and commands; they are not serialized into workbook JSON.

Treat callback and `getDocument()` results as immutable snapshots. Use `JSON.stringify` for persistence
and validate untrusted parsed data before passing it to a mounted component.

See the [SpreadsheetDocument API](/docs/api/interfaces/SpreadsheetDocument) or load compatible sparse JSON in the
[Legacy JSON Playground](/playground?mode=legacy-json).
