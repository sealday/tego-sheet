---
title: Frozen Panes and Layout
---

# Frozen Panes and Layout

Set `SheetData.freeze` to the A1 address of the first unfrozen cell. For example, `B2` keeps row 1 and
column A visible while the rest of the sheet scrolls. `A1` means that no rows or columns are frozen.

```tsx
import { useRef } from 'react';
import { TegoSheet, type SheetData, type SheetOptions, type TegoSheetHandle } from 'tego-sheet';
import 'tego-sheet/styles.css';

const sheet: SheetData = {
  name: 'Schedule',
  freeze: 'B2',
  rows: { len: 200 },
  cols: { len: 24 },
};

const options: SheetOptions = {
  rows: { defaultHeight: 28 },
  columns: { defaultWidth: 120, minimumWidth: 48 },
  rowHeaderWidth: 52,
};

export function FrozenSchedule() {
  const ref = useRef<TegoSheetHandle>(null);

  return (
    <>
      <button type="button" onClick={() => ref.current?.recalculateLayout()}>
        Recalculate after panel resize
      </button>
      <div style={{ height: 'min(70vh, 720px)', minWidth: 0 }}>
        <TegoSheet ref={ref} defaultValue={[sheet]} options={options} />
      </div>
    </>
  );
}
```

Selections use inclusive, zero-based ranges and keep one active cell inside the range. Frozen panes do
not change those coordinates; they only divide rendering and scrolling. Row/column defaults, header
width, default style, and focus settings are captured at mount, so remount the component to change
those options.

The component fills its parent. Give that parent a real height and `min-width: 0` when it participates
in flex or grid layout. Normal window and observed container changes are recalculated automatically;
call `recalculateLayout()` after an application-owned transition that does not produce a timely size
notification.

Review the [Styling and Sizing guide](../getting-started/styling-and-sizing), then see the
[SheetData](/docs/api/interfaces/SheetData) and [SheetOptions](/docs/api/interfaces/SheetOptions) APIs.
