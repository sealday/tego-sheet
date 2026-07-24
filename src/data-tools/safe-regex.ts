import { DataTransformError } from './errors';

export interface SafeRegexLimits {
  readonly maximumPatternLength: number;
  readonly maximumInputLength: number;
  readonly maximumSteps: number;
  readonly maximumMilliseconds: number;
}

function quantifierEnd(source: string, index: number): number | undefined {
  const token = source[index];
  if (token === '*' || token === '+' || token === '?') return index + 1;
  if (token !== '{') return undefined;
  const match = /^\{\d+(?:,\d*)?\}/u.exec(source.slice(index));
  return match === null ? undefined : index + match[0].length;
}

function hasUnsafeNestedQuantifier(source: string): boolean {
  const groups: { quantified: boolean; alternation: boolean }[] = [];
  let inClass = false;
  for (let index = 0; index < source.length; index += 1) {
    const token = source[index];
    if (token === '\\') {
      index += 1;
      continue;
    }
    if (token === '[') {
      inClass = true;
      continue;
    }
    if (token === ']' && inClass) {
      inClass = false;
      continue;
    }
    if (inClass) continue;
    if (token === '(') {
      groups.push({ quantified: false, alternation: false });
      continue;
    }
    if (token === '|') {
      const current = groups.at(-1);
      if (current !== undefined) current.alternation = true;
      continue;
    }
    if (token === ')') {
      const group = groups.pop();
      if (group === undefined) continue;
      const end = quantifierEnd(source, index + 1);
      if (end === undefined) continue;
      if (group.quantified || group.alternation) return true;
      const parent = groups.at(-1);
      if (parent !== undefined) parent.quantified = true;
      continue;
    }
    const end = quantifierEnd(source, index);
    if (end !== undefined) {
      const current = groups.at(-1);
      if (current !== undefined) current.quantified = true;
      index = end - 1;
      if (source[index + 1] === '?') index += 1;
    }
  }
  return false;
}

function assertSafePattern(source: string, limits: SafeRegexLimits): void {
  if (source.length > limits.maximumPatternLength) {
    throw new DataTransformError('REPLACE_PATTERN_INVALID', 'Replacement pattern is too long');
  }
  if (/\\(?:[1-9]\d*|k<[^>]+>)/u.test(source)) {
    throw new DataTransformError(
      'REPLACE_PATTERN_INVALID',
      'Replacement patterns cannot contain backreferences',
    );
  }
  if (/\(\?(?:[=!]|<[=!])/u.test(source)) {
    throw new DataTransformError(
      'REPLACE_PATTERN_INVALID',
      'Replacement patterns cannot contain lookaround assertions',
    );
  }
  if (hasUnsafeNestedQuantifier(source)) {
    throw new DataTransformError(
      'REPLACE_PATTERN_INVALID',
      'Replacement pattern contains an unsafe nested or ambiguous quantifier',
    );
  }
}

/** One bounded regex instance scoped to a single transform preview. */
export interface SafeRegexBudget {
  readonly pattern: RegExp;
  replace(value: string, replacement: string): string;
}

export function createSafeRegexBudget(source: string, limits: SafeRegexLimits): SafeRegexBudget {
  assertSafePattern(source, limits);
  let pattern: RegExp;
  try {
    pattern = new RegExp(source, 'gu');
  } catch (cause) {
    throw new DataTransformError(
      'REPLACE_PATTERN_INVALID',
      cause instanceof Error ? cause.message : 'Replacement pattern is invalid',
    );
  }
  let steps = 0;
  let elapsed = 0;
  return {
    pattern,
    replace(value, replacement) {
      if (value.length > limits.maximumInputLength) {
        throw new DataTransformError(
          'REPLACE_BUDGET_EXCEEDED',
          'Replacement input exceeds the configured length budget',
        );
      }
      steps += Math.max(1, source.length) * Math.max(1, value.length);
      if (steps > limits.maximumSteps) {
        throw new DataTransformError(
          'REPLACE_BUDGET_EXCEEDED',
          'Replacement exceeds the configured regex step budget',
        );
      }
      pattern.lastIndex = 0;
      const started = performance.now();
      const result = value.replace(pattern, replacement);
      elapsed += performance.now() - started;
      if (elapsed > limits.maximumMilliseconds) {
        throw new DataTransformError(
          'REPLACE_BUDGET_EXCEEDED',
          'Replacement exceeds the configured regex time budget',
        );
      }
      return result;
    },
  };
}
