---
title: Validation and Filtering
---

# Validation and Filtering

Validation and auto-filter state are part of the serialized workbook. Use `ValidationData` and
`AutoFilterData` to prepare compatible data, then call `TegoSheetHandle.validate()` when the
application needs a complete validation report.

```tsx
import { useRef } from 'react';
import {
  TegoSheet,
  type AutoFilterData,
  type TegoSheetHandle,
  type ValidationData,
  type WorkbookData,
} from 'tego-sheet';
import 'tego-sheet/styles.css';

const amountRule: ValidationData = {
  refs: ['B2:B100'],
  mode: 'cell',
  type: 'number',
  required: true,
  operator: 'gte',
  value: 0,
};

const categoryFilter: AutoFilterData = {
  ref: 'A1:B100',
  filters: [{ ci: 0, operator: 'in', value: ['Hosting', 'Support'] }],
  sort: { ci: 1, order: 'desc' },
};

const workbook: WorkbookData = [
  {
    name: 'Expenses',
    rows: {
      len: 100,
      0: { cells: { 0: { text: 'Category' }, 1: { text: 'Amount' } } },
      1: { cells: { 0: { text: 'Hosting' }, 1: { text: '29' } } },
    },
    validations: [amountRule],
    autofilter: categoryFilter,
  },
];

export function ValidatedSheet() {
  const ref = useRef<TegoSheetHandle>(null);

  return (
    <>
      <button type="button" onClick={() => console.log(ref.current?.validate())}>
        Validate workbook
      </button>
      <div style={{ height: 480 }}>
        <TegoSheet ref={ref} defaultValue={workbook} />
      </div>
    </>
  );
}
```

Validation references are A1 cells or ranges. `validate()` checks every configured sheet and returns a
`ValidationResult` containing zero-based addresses for all failures. A filter's `ref` covers its header
and data rows; `ci` is an absolute zero-based worksheet column index, not an offset within the range.
Filtering changes row visibility, and sorting commits a workbook mutation.

Custom toolbars can create the same state through typed `ToolbarAction` values such as
`set-validation`, `remove-validation`, `set-filter`, `clear-filter`, and `sort`.

See [ValidationData](/docs/api/interfaces/ValidationData),
[AutoFilterData](/docs/api/interfaces/AutoFilterData), and
[TegoSheetHandle](/docs/api/interfaces/TegoSheetHandle).
