import type { JsonObject, JsonValue } from '../../core/types/json';
import type { BuiltInCellTypeDefinition } from '../kernel/capabilities';
import { extensionDiagnostic, ExtensionKernelError } from '../kernel/manifest';

export type CheckboxCellValue = JsonObject & {
  readonly checked: boolean;
  readonly label?: string;
};

function isRecord(value: JsonValue): value is Readonly<Record<string, JsonValue>> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function validate(value: JsonValue): value is CheckboxCellValue {
  if (!isRecord(value)) return false;
  const keys = Object.keys(value);
  return (
    keys.every((key) => key === 'checked' || key === 'label') &&
    typeof value.checked === 'boolean' &&
    (value.label === undefined || typeof value.label === 'string')
  );
}

function invalid(value: JsonValue, fromVersion: number): never {
  throw new ExtensionKernelError([
    extensionDiagnostic(
      'CELL_TYPE_VALUE_INVALID',
      'migrate',
      `Checkbox value is invalid for schema version ${fromVersion}`,
      { id: 'checkbox', kind: 'cell-type' },
      { fromVersion, value },
    ),
  ]);
}

/** Built-in checkbox value semantics. */
export const checkboxCellType: BuiltInCellTypeDefinition<CheckboxCellValue> = Object.freeze({
  id: 'checkbox',
  schemaVersion: 1,
  validate,
  migrate(value: JsonValue, fromVersion: number): CheckboxCellValue {
    if (fromVersion === 0) {
      if (typeof value !== 'boolean') return invalid(value, fromVersion);
      return Object.freeze({ checked: value });
    }
    if (fromVersion !== 1 || !validate(value)) return invalid(value, fromVersion);
    return Object.freeze({
      checked: value.checked,
      ...(value.label === undefined ? {} : { label: value.label }),
    });
  },
  describe(value: CheckboxCellValue) {
    const state = value.checked ? 'checked' : 'Not checked';
    return Object.freeze({
      formattedText: value.label ?? (value.checked ? 'TRUE' : 'FALSE'),
      accessibilityLabel:
        value.label === undefined ? state : `${value.label}: ${state.toLowerCase()}`,
      role: 'checkbox' as const,
      checked: value.checked,
    });
  },
  toFormulaScalar(value: CheckboxCellValue) {
    return value.checked;
  },
});
