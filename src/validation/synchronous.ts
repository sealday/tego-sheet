import type { ValidationComparison, ValidationRequest, ValidationResult } from './model';

/** @internal Builds the stable reject/warn result shared by sync and async validation. */
export function validationFailure(
  request: ValidationRequest,
  status: 'rejected' | 'warning',
): ValidationResult {
  return {
    status,
    code: 'VALIDATION_REJECTED',
    diagnostics: [{ code: 'VALIDATION_REJECTED', ruleId: request.rule.id }],
  };
}

function compare<Value extends number | string>(
  candidate: Value,
  predicate: ValidationComparison<Value>,
): boolean {
  switch (predicate.operator) {
    case 'between':
      return candidate >= predicate.minimum && candidate <= predicate.maximum;
    case 'notBetween':
      return candidate < predicate.minimum || candidate > predicate.maximum;
    case 'equal':
      return candidate === predicate.value;
    case 'notEqual':
      return candidate !== predicate.value;
    case 'greaterThan':
      return candidate > predicate.value;
    case 'lessThan':
      return candidate < predicate.value;
    case 'greaterThanOrEqual':
      return candidate >= predicate.value;
    case 'lessThanOrEqual':
      return candidate <= predicate.value;
  }
}

function mapComparison<Source extends number | string, Target extends number | string>(
  predicate: ValidationComparison<Source>,
  convert: (value: Source) => Target | undefined,
): ValidationComparison<Target> | undefined {
  if (predicate.operator === 'between' || predicate.operator === 'notBetween') {
    const minimum = convert(predicate.minimum);
    const maximum = convert(predicate.maximum);
    return minimum === undefined || maximum === undefined
      ? undefined
      : { operator: predicate.operator, minimum, maximum };
  }
  if (!('value' in predicate)) return undefined;
  const value = convert(predicate.value);
  return value === undefined ? undefined : { operator: predicate.operator, value };
}

function calendarDate(value: string): number | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/u.exec(value);
  if (match === null) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const timestamp = Date.UTC(year, month - 1, day);
  const parsed = new Date(timestamp);
  return parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day
    ? timestamp
    : undefined;
}

function localTime(value: string): number | undefined {
  const match = /^(\d{2}):(\d{2})(?::(\d{2}))?$/u.exec(value);
  if (match === null) return undefined;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] ?? 0);
  return hours <= 23 && minutes <= 59 && seconds <= 59
    ? hours * 3_600 + minutes * 60 + seconds
    : undefined;
}

function dateCandidate(request: ValidationRequest): number | undefined {
  if (request.value.type === 'number') return request.value.value;
  if (request.value.type !== 'string') return undefined;
  const timestamp = calendarDate(request.value.value);
  return timestamp === undefined ? undefined : timestamp / 86_400_000 + 25_569;
}

function timeCandidate(request: ValidationRequest): number | undefined {
  if (request.value.type === 'number') {
    return request.value.value >= 0 && request.value.value < 1
      ? request.value.value * 86_400
      : undefined;
  }
  return request.value.type === 'string' ? localTime(request.value.value) : undefined;
}

function textLength(request: ValidationRequest): number | undefined {
  switch (request.value.type) {
    case 'string':
      return request.value.value.length;
    case 'number':
    case 'boolean':
      return String(request.value.value).length;
    case 'blank':
      return 0;
    case 'array':
    case 'error':
      return undefined;
  }
}

/** @internal Evaluates only the static portion shared with the async engine. */
export function staticRuleAccepted(request: ValidationRequest): boolean | undefined {
  const { rule } = request;
  switch (rule.type) {
    case 'whole':
      return (
        request.value.type === 'number' &&
        Number.isInteger(request.value.value) &&
        compare(request.value.value, rule.predicate)
      );
    case 'decimal':
    case 'number':
      return request.value.type === 'number' && compare(request.value.value, rule.predicate);
    case 'date': {
      const candidate = dateCandidate(request);
      if (candidate === undefined) return false;
      const predicate = mapComparison(rule.predicate, (value) =>
        dateCandidate({ ...request, value: { type: 'string', value } }),
      );
      return predicate !== undefined && compare(candidate, predicate);
    }
    case 'time': {
      const candidate = timeCandidate(request);
      if (candidate === undefined) return false;
      const predicate = mapComparison(rule.predicate, (value) =>
        timeCandidate({ ...request, value: { type: 'string', value } }),
      );
      return predicate !== undefined && compare(candidate, predicate);
    }
    case 'text-length': {
      const candidate = textLength(request);
      return candidate !== undefined && compare(candidate, rule.predicate);
    }
    case 'list':
    case 'custom-formula':
      return undefined;
  }
}

/** Synchronous result used by the document mutation boundary. */
export type SynchronousValidationResult =
  | ValidationResult
  | { readonly status: 'async-required'; readonly code: 'ASYNC_REQUIRED' };

/** Evaluates rules that require no host resolver or formula capability. */
export function validateValidationRequestSync(
  request: ValidationRequest,
): SynchronousValidationResult {
  if (request.value.type === 'blank' && request.rule.allowBlank) {
    return { status: 'accepted', diagnostics: [] };
  }
  if (
    request.rule.type === 'custom-formula' ||
    (request.rule.type === 'list' && request.rule.predicate.source.type === 'resolver')
  ) {
    return { status: 'async-required', code: 'ASYNC_REQUIRED' };
  }
  let accepted = staticRuleAccepted(request);
  if (request.rule.type === 'list') {
    const source = request.rule.predicate.source;
    if (source.type !== 'static') return { status: 'async-required', code: 'ASYNC_REQUIRED' };
    accepted = request.value.type === 'string' && new Set(source.values).has(request.value.value);
  }
  return accepted === true
    ? { status: 'accepted', diagnostics: [] }
    : validationFailure(request, request.rule.behavior === 'warn' ? 'warning' : 'rejected');
}
