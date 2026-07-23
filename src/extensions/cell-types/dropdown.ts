import type { JsonObject, JsonValue } from '../../core/types/json';
import type { BuiltInCellTypeDefinition, CellTypeScalar } from '../kernel/capabilities';
import { extensionDiagnostic, ExtensionKernelError } from '../kernel/manifest';

export type DropdownCellValue = JsonObject & {
  readonly value: CellTypeScalar;
  readonly label?: string;
};

function isScalar(value: JsonValue): value is CellTypeScalar {
  return (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean' ||
    (typeof value === 'number' && Number.isFinite(value))
  );
}

function isRecord(value: JsonValue): value is Readonly<Record<string, JsonValue>> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function validate(value: JsonValue): value is DropdownCellValue {
  if (!isRecord(value)) return false;
  const keys = Object.keys(value);
  return (
    keys.every((key) => key === 'value' || key === 'label') &&
    Object.hasOwn(value, 'value') &&
    isScalar(value.value!) &&
    (!('label' in value) || (Object.hasOwn(value, 'label') && typeof value.label === 'string'))
  );
}

function invalid(value: JsonValue, fromVersion: number): never {
  throw new ExtensionKernelError([
    extensionDiagnostic(
      'CELL_TYPE_VALUE_INVALID',
      'migrate',
      `Dropdown value is invalid for schema version ${fromVersion}`,
      { id: 'dropdown', kind: 'cell-type' },
      { fromVersion, value },
    ),
  ]);
}

/** Built-in dropdown value semantics. */
export const dropdownCellType: BuiltInCellTypeDefinition<DropdownCellValue> = Object.freeze({
  id: 'dropdown',
  schemaVersion: 1,
  validate,
  migrate(value: JsonValue, fromVersion: number): DropdownCellValue {
    if (fromVersion === 0) {
      if (!isScalar(value)) return invalid(value, fromVersion);
      return Object.freeze({ value });
    }
    if (fromVersion !== 1 || !validate(value)) return invalid(value, fromVersion);
    return Object.freeze({
      value: value.value,
      ...(value.label === undefined ? {} : { label: value.label }),
    });
  },
  describe(value: DropdownCellValue) {
    const formattedText = value.label ?? (value.value === null ? '' : String(value.value));
    return Object.freeze({
      formattedText,
      accessibilityLabel: formattedText === '' ? 'Blank' : formattedText,
      role: 'combobox' as const,
    });
  },
  toFormulaScalar(value: DropdownCellValue) {
    return value.value;
  },
});
