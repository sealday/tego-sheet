---
title: Styling and Sizing
---

# Styling and Sizing

`TegoSheet` fills its parent with `width: 100%` and `height: 100%`. Give an ancestor an explicit
height; otherwise the browser cannot resolve that percentage to a useful grid viewport. The component
has a 240px minimum height, but that minimum is not a replacement for a deliberate application layout.

```tsx
import { TegoSheet } from 'tego-sheet';
import 'tego-sheet/styles.css';
import './sheet-panel.css';

export function SheetPanel() {
  return (
    <section className="sheet-panel" aria-label="Inventory workbook">
      <TegoSheet defaultValue={[{ name: 'Inventory' }]} />
    </section>
  );
}
```

```css title="sheet-panel.css"
.sheet-panel {
  width: 100%;
  height: clamp(360px, 70vh, 760px);
  min-width: 0;
  min-height: 0;
}

@media (max-width: 640px) {
  .sheet-panel {
    height: 65vh;
  }
}
```

The CSS is theme-neutral: it controls only layout and leaves the package stylesheet responsible for
spreadsheet chrome. Use the `className` and `style` props when an instance needs an additional hook.
If a container changes size without a window resize—for example, after a split pane moves—call
`TegoSheetHandle.recalculateLayout()` after the layout settles.

See the [TegoSheet props API](/docs/api/interfaces/TegoSheetProps) and the
[layout guide](../guides/frozen-panes-and-layout).
