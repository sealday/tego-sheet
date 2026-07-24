---
sidebar_label: Template API
---

# Template compilation and generation

Template APIs compile a constrained workbook template and structured input into an immutable
generated document. That result owns pagination, resources, and the print display list shared by
browser print and output adapters.

```ts
import { compileSpreadsheetTemplate, renderSpreadsheetTemplate } from 'tego-sheet';

const compiled = compileSpreadsheetTemplate(template);
const generated = await renderSpreadsheetTemplate(compiled, input, {
  signal: abortController.signal,
});
```

Bindings support scalar values, conditions, row/column/range/page/sheet repeats, and
subtemplates. Expressions cannot execute JavaScript or access ambient globals. Compilation and
rendering enforce node, clone, page, resource, byte, and time budgets and fail atomically with
structured diagnostics.

Public template modules live in `tego-sheet/sdk`. They run the fixed
recognize → transform → resolve → layout → paint pipeline and may emit only validated,
renderer-neutral `PrintDisplayCommand` values. A module cannot mutate the workbook or bypass
resource resolution.
