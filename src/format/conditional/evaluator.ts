import {
  BUILTIN_FORMULA_COMPATIBILITY,
  parseFormula,
  type FormulaAst,
  type FormulaValue,
} from '../../formula';
import type {
  ConditionalEvaluationInput,
  ConditionalEvaluationResult,
  ConditionalExpression,
  ConditionalRule,
} from './model';

/** Stable conditional-format failure. */
export class ConditionalFormatError extends Error {
  /** Creates a stable conditional-format error. */
  constructor(
    /** Machine-readable failure category. */
    readonly code: 'INVALID_CONDITIONAL_EXPRESSION' | 'CONDITIONAL_RANGE_TOO_LARGE',
    message: string,
  ) {
    super(message);
    this.name = 'ConditionalFormatError';
  }
}

/** Resource limits enforced before conditional evaluation. */
export interface ConditionalFormatLimits {
  /** Maximum rules accepted per evaluation. */
  readonly maxRules: number;
  /** Maximum aggregate target area accepted per evaluation. */
  readonly maxCells: number;
}

function contains(rule: ConditionalRule, address: ConditionalEvaluationInput['address']): boolean {
  return rule.ranges.some(
    (range) =>
      range.sheetId === address.sheetId &&
      address.row >= range.start.row &&
      address.row <= range.end.row &&
      address.column >= range.start.column &&
      address.column <= range.end.column,
  );
}

function scalar(value: FormulaValue): number | string | boolean | undefined {
  if (value.type === 'blank' || value.type === 'array' || value.type === 'error') return undefined;
  return value.value;
}

function compare(
  left: number | string | boolean | undefined,
  operator: Extract<ConditionalExpression, { type: 'cell-is' }>['operator'],
  right: number | string,
  right2?: number | string,
): boolean {
  if (left === undefined) return false;
  const comparison =
    typeof left === 'number' && typeof right === 'number'
      ? left - right
      : String(left).localeCompare(String(right), 'en-US');
  if (operator === 'equal') return comparison === 0;
  if (operator === 'notEqual') return comparison !== 0;
  if (operator === 'between' || operator === 'notBetween') {
    if (right2 === undefined) return false;
    const upperComparison =
      typeof left === 'number' && typeof right2 === 'number'
        ? left - right2
        : String(left).localeCompare(String(right2), 'en-US');
    const between = comparison >= 0 && upperComparison <= 0;
    return operator === 'between' ? between : !between;
  }
  if (operator === 'greaterThan') return comparison > 0;
  if (operator === 'greaterThanOrEqual') return comparison >= 0;
  if (operator === 'lessThan') return comparison < 0;
  return comparison <= 0;
}

const allowedFunctions = new Set(BUILTIN_FORMULA_COMPATIBILITY.map(({ name }) => name));

function validateFormula(ast: FormulaAst): void {
  if (ast.kind === 'call') {
    if (!allowedFunctions.has(ast.name)) {
      throw new ConditionalFormatError(
        'INVALID_CONDITIONAL_EXPRESSION',
        `Conditional formula function ${ast.name} is unsupported`,
      );
    }
    for (const argument of ast.arguments) validateFormula(argument);
  } else if (ast.kind === 'binary') {
    validateFormula(ast.left);
    validateFormula(ast.right);
  } else if (ast.kind === 'unary') {
    validateFormula(ast.operand);
  }
}

function matches(condition: ConditionalExpression, input: ConditionalEvaluationInput): boolean {
  if (condition.type === 'blank') return input.value.type === 'blank';
  if (condition.type === 'not-blank') return input.value.type !== 'blank';
  if (condition.type === 'text-contains') return input.text.includes(condition.text);
  if (condition.type === 'cell-is') {
    return compare(scalar(input.value), condition.operator, condition.value, condition.value2);
  }
  try {
    const ast = parseFormula(condition.source);
    validateFormula(ast);
    if (ast.kind === 'boolean') return ast.value;
    if (ast.kind === 'number') return ast.value !== 0;
    if (ast.kind === 'reference') {
      return Boolean(
        scalar(
          input.lookup?.({
            sheetId: ast.reference.sheetId ?? input.address.sheetId,
            row: ast.reference.row,
            column: ast.reference.column,
          }) ?? { type: 'blank' },
        ),
      );
    }
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      'Conditional formula is outside the synchronous boolean subset',
    );
  } catch (error) {
    if (error instanceof ConditionalFormatError) throw error;
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      error instanceof Error ? error.message : String(error),
    );
  }
}

/** Creates an isolated pure conditional-format evaluator. */
export function createConditionalFormatEvaluator(limits: ConditionalFormatLimits): {
  evaluate(input: ConditionalEvaluationInput): ConditionalEvaluationResult;
} {
  return {
    evaluate(input) {
      if (
        input.rules.length > limits.maxRules ||
        input.rules.reduce(
          (count, rule) =>
            count +
            rule.ranges.reduce(
              (cells, range) =>
                cells +
                (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1),
              0,
            ),
          0,
        ) > limits.maxCells
      ) {
        throw new ConditionalFormatError(
          'CONDITIONAL_RANGE_TOO_LARGE',
          'Conditional formatting exceeds configured limits',
        );
      }
      const patch: Record<string, import('../../core/types/json').JsonValue> = {};
      const matched: string[] = [];
      const diagnostics: ConditionalEvaluationResult['diagnostics'][number][] = [];
      for (const rule of [...input.rules].sort(
        (left, right) => left.priority - right.priority || left.id.localeCompare(right.id),
      )) {
        if (!contains(rule, input.address) || !matches(rule.condition, input)) continue;
        matched.push(rule.id);
        if (rule.effect.type === 'style') Object.assign(patch, rule.effect.patch);
        else diagnostics.push({ code: 'UNSUPPORTED_FORMAT_FEATURE', ruleId: rule.id });
        if (rule.stopIfTrue) break;
      }
      return {
        matchedRuleIds: Object.freeze(matched),
        stylePatch: Object.freeze(patch),
        diagnostics: Object.freeze(diagnostics),
      };
    },
  };
}
