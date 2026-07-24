import { DataTransformError } from './errors';

export interface SafeRegexLimits {
  readonly maximumPatternLength: number;
  readonly maximumInputLength: number;
  readonly maximumOutputLength: number;
  readonly maximumSteps: number;
  readonly maximumMilliseconds: number;
}

function assertLinearPattern(source: string): void {
  let inClass = false;
  let quantifiers = 0;
  for (let index = 0; index < source.length; index += 1) {
    const token = source[index];
    if (token === '\\') {
      if (source[index + 1] === 'p' || source[index + 1] === 'P') {
        throw new DataTransformError(
          'REPLACE_PATTERN_INVALID',
          'Replacement patterns cannot contain Unicode property escapes',
        );
      }
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
    if (token === '|') {
      throw new DataTransformError(
        'REPLACE_PATTERN_INVALID',
        'Replacement patterns cannot contain alternation',
      );
    }
    if (token === '(' && source[index + 1] === '?') {
      throw new DataTransformError(
        'REPLACE_PATTERN_INVALID',
        'Replacement patterns cannot contain special groups',
      );
    }
    if (token === '*' || token === '+') {
      throw new DataTransformError(
        'REPLACE_PATTERN_INVALID',
        'Replacement patterns cannot contain unbounded quantifiers',
      );
    }
    if (token !== '?' && token !== '{') continue;
    if (source[index - 1] === ')') {
      throw new DataTransformError(
        'REPLACE_PATTERN_INVALID',
        'Replacement patterns cannot quantify groups',
      );
    }
    let end = index + 1;
    if (token === '{') {
      const match = /^\{(\d+)(?:,(\d+))?\}/u.exec(source.slice(index));
      if (
        match === null ||
        Number(match[1]) > 100 ||
        (match[2] !== undefined && Number(match[2]) > 100)
      ) {
        throw new DataTransformError(
          'REPLACE_PATTERN_INVALID',
          'Replacement patterns require a finite bounded quantifier no greater than 100',
        );
      }
      end = index + match[0].length;
    }
    quantifiers += 1;
    if (quantifiers > 1 || source[end] === '?') {
      throw new DataTransformError(
        'REPLACE_PATTERN_INVALID',
        'Replacement patterns can contain at most one non-lazy bounded quantifier',
      );
    }
    index = end - 1;
  }
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
  assertLinearPattern(source);
}

/** One bounded regex instance scoped to a single transform preview. */
export interface SafeRegexBudget {
  readonly pattern: RegExp;
  replace(value: string, replacement: string): string;
}

function advanceUnicodeIndex(value: string, index: number): number {
  const codePoint = value.codePointAt(index);
  return index + (codePoint !== undefined && codePoint > 0xffff ? 2 : 1);
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
  let outputCodeUnits = 0;
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
      let outputLength = value.length;
      let matched = false;
      let match: RegExpExecArray | null;
      while ((match = pattern.exec(value)) !== null) {
        matched = true;
        outputLength += replacement.length - match[0].length;
        if (
          !Number.isSafeInteger(outputLength) ||
          outputLength > limits.maximumInputLength ||
          outputCodeUnits > limits.maximumOutputLength - outputLength
        ) {
          throw new DataTransformError(
            'REPLACE_BUDGET_EXCEEDED',
            'Replacement output exceeds the configured length budget',
          );
        }
        if (match[0] === '') {
          pattern.lastIndex = advanceUnicodeIndex(value, pattern.lastIndex);
        }
      }
      if (!matched) {
        elapsed += performance.now() - started;
        if (elapsed > limits.maximumMilliseconds) {
          throw new DataTransformError(
            'REPLACE_BUDGET_EXCEEDED',
            'Replacement exceeds the configured regex time budget',
          );
        }
        return value;
      }
      outputCodeUnits += outputLength;
      pattern.lastIndex = 0;
      const result = value.replace(pattern, () => replacement);
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
