import { describe, expect, it } from 'vitest';
import type { JsonValue } from '../../../src/core/types/json';
import { checkboxCellType } from '../../../src/extensions/cell-types/checkbox';
import { dropdownCellType } from '../../../src/extensions/cell-types/dropdown';

const environment = { locale: 'en-US', timeZone: 'UTC' };

describe('checkbox cell type', () => {
  it('validates only the canonical JSON value', () => {
    expect(checkboxCellType.validate({ checked: false })).toBe(true);
    expect(checkboxCellType.validate({ checked: true, label: 'Approved' })).toBe(true);
    expect(checkboxCellType.validate(true)).toBe(false);
    expect(checkboxCellType.validate({ checked: 1 })).toBe(false);
    expect(checkboxCellType.validate({ checked: true, extra: true })).toBe(false);
  });

  it('migrates the legacy boolean and rejects unknown or invalid schemas atomically', () => {
    expect(checkboxCellType.migrate?.(false, 0)).toEqual({ checked: false });
    expect(checkboxCellType.migrate?.({ checked: true, label: 'Approved' }, 1)).toEqual({
      checked: true,
      label: 'Approved',
    });
    expect(() => checkboxCellType.migrate?.('true', 0)).toThrowError(
      expect.objectContaining({ code: 'CELL_TYPE_VALUE_INVALID' }),
    );
    expect(() => checkboxCellType.migrate?.(true, 9)).toThrowError(
      expect.objectContaining({ code: 'CELL_TYPE_VALUE_INVALID' }),
    );
  });

  it('derives formatted, accessibility, and formula semantics from one value', () => {
    const checked = { checked: true, label: 'Approved' } as const;
    expect(checkboxCellType.describe(checked, environment)).toEqual({
      formattedText: 'Approved',
      accessibilityLabel: 'Approved: checked',
      role: 'checkbox',
      checked: true,
    });
    expect(checkboxCellType.toFormulaScalar(checked)).toBe(true);

    expect(checkboxCellType.describe({ checked: false }, environment)).toEqual({
      formattedText: 'FALSE',
      accessibilityLabel: 'Not checked',
      role: 'checkbox',
      checked: false,
    });
  });
});

describe('dropdown cell type', () => {
  it('validates canonical scalar values and an optional display label', () => {
    const values: JsonValue[] = [
      { value: null },
      { value: '' },
      { value: 0 },
      { value: false },
      { value: 'approved', label: 'Approved' },
    ];
    for (const value of values) {
      expect(dropdownCellType.validate(value)).toBe(true);
    }
    expect(dropdownCellType.validate({ value: [] })).toBe(false);
    expect(dropdownCellType.validate({ value: 'a', label: 1 })).toBe(false);
    expect(dropdownCellType.validate({ value: 'a', extra: true })).toBe(false);
  });

  it('migrates a legacy scalar and rejects arrays and unknown schemas', () => {
    expect(dropdownCellType.migrate?.(0, 0)).toEqual({ value: 0 });
    expect(dropdownCellType.migrate?.({ value: false }, 1)).toEqual({ value: false });
    expect(() => dropdownCellType.migrate?.([], 0)).toThrowError(
      expect.objectContaining({ code: 'CELL_TYPE_VALUE_INVALID' }),
    );
    expect(() => dropdownCellType.migrate?.('value', 9)).toThrowError(
      expect.objectContaining({ code: 'CELL_TYPE_VALUE_INVALID' }),
    );
  });

  it('preserves falsy formula scalars and uses labels only for presentation', () => {
    expect(dropdownCellType.describe({ value: 0, label: 'None' }, environment)).toEqual({
      formattedText: 'None',
      accessibilityLabel: 'None',
      role: 'combobox',
    });
    expect(dropdownCellType.toFormulaScalar({ value: 0, label: 'None' })).toBe(0);
    expect(dropdownCellType.describe({ value: null }, environment)).toEqual({
      formattedText: '',
      accessibilityLabel: 'Blank',
      role: 'combobox',
    });
    expect(dropdownCellType.toFormulaScalar({ value: false })).toBe(false);
  });
});
