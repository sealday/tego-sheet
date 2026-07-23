import type {
  DocumentCellAddress,
  JsonValue,
  SpreadsheetDocument,
} from '../document/model/document';
import type { FormulaValue } from '../formula';
import type { ValidationRequest, ValidationRule } from './model';

function record(value: JsonValue): Readonly<Record<string, JsonValue>> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Readonly<Record<string, JsonValue>>)
    : undefined;
}

function validationRule(value: JsonValue): ValidationRule | undefined {
  const item = record(value);
  if (
    item === undefined ||
    typeof item.id !== 'string' ||
    (item.type !== 'number' && item.type !== 'list') ||
    (item.behavior !== 'reject' && item.behavior !== 'warn') ||
    typeof item.allowBlank !== 'boolean'
  ) {
    return undefined;
  }
  const predicate = record(item.predicate);
  if (predicate === undefined) return undefined;
  if (
    item.type === 'number' &&
    predicate.operator === 'between' &&
    typeof predicate.minimum === 'number' &&
    Number.isFinite(predicate.minimum) &&
    typeof predicate.maximum === 'number' &&
    Number.isFinite(predicate.maximum)
  ) {
    return {
      id: item.id,
      type: 'number',
      predicate: {
        operator: 'between',
        minimum: predicate.minimum,
        maximum: predicate.maximum,
      },
      behavior: item.behavior,
      allowBlank: item.allowBlank,
    };
  }
  const source = record(predicate.source);
  if (
    item.type === 'list' &&
    source !== undefined &&
    ((source.type === 'static' &&
      Array.isArray(source.values) &&
      source.values.every((entry) => typeof entry === 'string')) ||
      (source.type === 'resolver' && typeof source.id === 'string'))
  ) {
    return {
      id: item.id,
      type: 'list',
      predicate: {
        source:
          source.type === 'static'
            ? { type: 'static', values: source.values as string[] }
            : { type: 'resolver', id: source.id as string },
      },
      behavior: item.behavior,
      allowBlank: item.allowBlank,
    };
  }
  return undefined;
}

function proposedValue(text: string, rule: ValidationRule): FormulaValue {
  if (text.length === 0) return { type: 'blank' };
  if (rule.type === 'number') {
    const value = Number(text);
    if (Number.isFinite(value)) return { type: 'number', value };
  }
  return { type: 'string', value: text };
}

/** @internal Resolves a cell-owned typed validation rule from one immutable document snapshot. */
export function documentValidationRequest(
  document: SpreadsheetDocument,
  address: DocumentCellAddress,
  text: string,
  signal?: AbortSignal,
): ValidationRequest | undefined {
  const sheet = document.workbook.sheets.find(({ id }) => id === address.sheetId);
  const cell = sheet?.cells.find(
    (entry) => entry.row === address.row && entry.column === address.column,
  )?.cell;
  if (cell?.validationId === undefined) return undefined;
  const entry = document.workbook.validations.find(({ id }) => id === cell.validationId);
  if (entry === undefined) return undefined;
  const rule = validationRule(entry.value);
  if (rule === undefined) return undefined;
  return {
    address,
    value: proposedValue(text, rule),
    rule,
    ...(signal === undefined ? {} : { signal }),
  };
}
