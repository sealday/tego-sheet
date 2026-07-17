---
title: Printing
---

# Printing

Printing is an imperative command on `TegoSheetHandle`. It renders the active worksheet for A4 portrait
output and opens the browser print flow. Call it directly from a user gesture so browser policies can
allow the dialog.

```tsx
import { useRef } from 'react';
import { TegoSheet, type TegoSheetError, type TegoSheetHandle } from 'tego-sheet';
import 'tego-sheet/styles.css';

export function PrintableSheet() {
  const ref = useRef<TegoSheetHandle>(null);

  function handlePrintError(error: TegoSheetError) {
    if (error.code === 'PRINT_FAILED') alert(error.message);
  }

  return (
    <>
      <button type="button" onClick={() => ref.current?.print()}>
        Print active sheet
      </button>
      <div style={{ height: 480 }}>
        <TegoSheet ref={ref} defaultValue={[{ name: 'Report' }]} onError={handlePrintError} />
      </div>
    </>
  );
}
```

The default and custom toolbar can also invoke the supported `print` action. Printing remains available
in `readOnly` mode. A cell with `printable: false` omits its content while preserving grid geometry,
style, and merged-cell layout. Printing failures handled by the component arrive through `onError` as
`PRINT_FAILED`; invalid handle use can throw `TegoSheetException` synchronously.

See the [TegoSheetHandle API](/docs/api/interfaces/TegoSheetHandle) and
[TegoSheetError API](/docs/api/interfaces/TegoSheetError).
