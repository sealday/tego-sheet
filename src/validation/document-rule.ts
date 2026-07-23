import type {
  CellInput,
  DocumentCellAddress,
  JsonValue,
  SpreadsheetDocument,
} from '../document/model/document';
import type { FormulaValue } from '../formula';
import type {
  ValidationComparison,
  ValidationComparisonOperator,
  ValidationRequest,
  ValidationRule,
  ValidationRuleBase,
} from './model';

function record(value: JsonValue): Readonly<Record<string, JsonValue>> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Readonly<Record<string, JsonValue>>)
    : undefined;
}

function base(item: Readonly<Record<string, JsonValue>>): ValidationRuleBase | undefined {
  return typeof item.id === 'string' &&
    (item.behavior === 'reject' || item.behavior === 'warn') &&
    typeof item.allowBlank === 'boolean'
    ? { id: item.id, behavior: item.behavior, allowBlank: item.allowBlank }
    : undefined;
}

const scalarOperators = new Set<ValidationComparisonOperator>([
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
]);

function comparison<Value extends number | string>(
  value: JsonValue,
  isValue: (candidate: JsonValue) => candidate is Value,
): ValidationComparison<Value> | undefined {
  const predicate = record(value);
  if (predicate === undefined || typeof predicate.operator !== 'string') return undefined;
  if (predicate.operator === 'between' || predicate.operator === 'notBetween') {
    return isValue(predicate.minimum) && isValue(predicate.maximum)
      ? {
          operator: predicate.operator,
          minimum: predicate.minimum,
          maximum: predicate.maximum,
        }
      : undefined;
  }
  return scalarOperators.has(predicate.operator as ValidationComparisonOperator) &&
    isValue(predicate.value)
    ? {
        operator: predicate.operator as Exclude<
          ValidationComparisonOperator,
          'between' | 'notBetween'
        >,
        value: predicate.value,
      }
    : undefined;
}

function validationRule(value: JsonValue): ValidationRule | undefined {
  const item = record(value);
  if (item === undefined) return undefined;
  const shared = base(item);
  if (shared === undefined) return undefined;

  if (item.type === 'whole' || item.type === 'decimal' || item.type === 'number') {
    const predicate = comparison(
      item.predicate,
      (candidate): candidate is number =>
        typeof candidate === 'number' && Number.isFinite(candidate),
    );
    return predicate === undefined ? undefined : { ...shared, type: item.type, predicate };
  }
  if (item.type === 'date' || item.type === 'time') {
    const predicate = comparison(
      item.predicate,
      (candidate): candidate is string => typeof candidate === 'string',
    );
    return predicate === undefined ? undefined : { ...shared, type: item.type, predicate };
  }
  if (item.type === 'text-length') {
    const predicate = comparison(
      item.predicate,
      (candidate): candidate is number =>
        typeof candidate === 'number' && Number.isFinite(candidate),
    );
    return predicate === undefined ? undefined : { ...shared, type: item.type, predicate };
  }
  const predicate = record(item.predicate);
  if (predicate === undefined) return undefined;
  if (item.type === 'custom-formula') {
    return typeof predicate.formula === 'string'
      ? {
          ...shared,
          type: item.type,
          predicate: { formula: predicate.formula },
        }
      : undefined;
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
      ...shared,
      type: 'list',
      predicate: {
        source:
          source.type === 'static'
            ? { type: 'static', values: source.values as string[] }
            : { type: 'resolver', id: source.id as string },
      },
    };
  }
  return undefined;
}

function customScalar(text: string, input: CellInput): FormulaValue | undefined {
  if (input.type !== 'custom') return undefined;
  if (input.cellType === 'checkbox') {
    if (/^true$/iu.test(text)) return { type: 'boolean', value: true };
    if (/^false$/iu.test(text)) return { type: 'boolean', value: false };
    return undefined;
  }
  if (input.cellType !== 'dropdown') return undefined;
  const payload = record(input.value);
  const current = payload?.value;
  if (typeof current === 'number') {
    const value = Number(text);
    return Number.isFinite(value) ? { type: 'number', value } : undefined;
  }
  if (typeof current === 'boolean') {
    if (/^true$/iu.test(text)) return { type: 'boolean', value: true };
    if (/^false$/iu.test(text)) return { type: 'boolean', value: false };
    return undefined;
  }
  return current === null && text.length === 0
    ? { type: 'blank' }
    : { type: 'string', value: text };
}

function proposedValue(text: string, rule: ValidationRule, input: CellInput): FormulaValue {
  if (text.length === 0) return { type: 'blank' };
  const custom = customScalar(text, input);
  if (custom !== undefined) return custom;
  if (rule.type === 'whole' || rule.type === 'decimal' || rule.type === 'number') {
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
    value: proposedValue(text, rule, cell.input),
    rule,
    ...(signal === undefined ? {} : { signal }),
  };
}
