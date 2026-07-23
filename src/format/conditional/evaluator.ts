import {
  BUILTIN_FORMULA_COMPATIBILITY,
  createFormulaFunctionRegistry,
  parseFormula,
  type FormulaAst,
  type ScalarFormulaValue,
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
  /** Maximum characters accepted in one conditional formula. */
  readonly maxFormulaLength?: number;
  /** Maximum parsed AST nodes accepted in one conditional formula. */
  readonly maxAstNodes?: number;
  /** Maximum recursive evaluation steps allowed per formula match. */
  readonly maxEvaluationSteps?: number;
}

interface FormulaRuntime {
  readonly asts: Map<string, FormulaAst>;
  readonly maxFormulaLength: number;
  readonly maxAstNodes: number;
  readonly maxEvaluationSteps: number;
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
  operator: Extract<ConditionalExpression, { type: 'cell-is' | 'cell-is-formula' }>['operator'],
  right: number | string | boolean,
  right2?: number | string | boolean,
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

function numeric(value: ScalarFormulaValue): number | undefined {
  if (value.type === 'number') return value.value;
  if (value.type === 'boolean') return value.value ? 1 : 0;
  if (value.type === 'blank') return 0;
  if (value.type === 'string') {
    const converted = Number(value.value);
    return Number.isFinite(converted) ? converted : undefined;
  }
  return undefined;
}

function truthy(value: ScalarFormulaValue): boolean {
  if (value.type === 'blank' || value.type === 'error') return false;
  return value.type === 'boolean' ? value.value : Boolean(value.value);
}

function evaluateFormulaAst(
  ast: FormulaAst,
  input: ConditionalEvaluationInput,
  budget: { remaining: number },
): ScalarFormulaValue {
  budget.remaining -= 1;
  if (budget.remaining < 0) {
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      'Conditional formula exceeded its evaluation budget',
    );
  }
  if (ast.kind === 'number' || ast.kind === 'string' || ast.kind === 'boolean') {
    return { type: ast.kind, value: ast.value } as ScalarFormulaValue;
  }
  if (ast.kind === 'error') return { type: 'error', value: ast.value };
  if (ast.kind === 'reference') {
    const value = input.lookup?.({
      sheetId: ast.reference.sheetId ?? input.address.sheetId,
      row: ast.reference.row,
      column: ast.reference.column,
    }) ?? { type: 'blank' };
    return value.type === 'array' ? { type: 'error', value: '#VALUE!' } : value;
  }
  if (ast.kind === 'range') {
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      'Conditional formula ranges are not scalar expressions',
    );
  }
  if (ast.kind === 'unary') {
    const value = numeric(evaluateFormulaAst(ast.operand, input, budget));
    return value === undefined
      ? { type: 'error', value: '#VALUE!' }
      : { type: 'number', value: -value };
  }
  if (ast.kind === 'call') {
    const definition = createFormulaFunctionRegistry().resolve(ast.name);
    if (definition === undefined || definition.mode !== 'sync') {
      throw new ConditionalFormatError(
        'INVALID_CONDITIONAL_EXPRESSION',
        `Conditional formula function ${ast.name} is unsupported`,
      );
    }
    const value = definition.evaluate(
      ast.arguments.map((argument) => evaluateFormulaAst(argument, input, budget)),
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        now: 0,
      },
    );
    if (value instanceof Promise || value.type === 'array') {
      throw new ConditionalFormatError(
        'INVALID_CONDITIONAL_EXPRESSION',
        'Conditional formulas must resolve synchronously to a scalar',
      );
    }
    return value;
  }
  const left = evaluateFormulaAst(ast.left, input, budget);
  const right = evaluateFormulaAst(ast.right, input, budget);
  if (left.type === 'error') return left;
  if (right.type === 'error') return right;
  const leftValue = left.type === 'blank' ? '' : left.value;
  const rightValue = right.type === 'blank' ? '' : right.value;
  if (ast.operator === '&') {
    return { type: 'string', value: `${String(leftValue)}${String(rightValue)}` };
  }
  const leftNumber = numeric(left);
  const rightNumber = numeric(right);
  if (['=', '==', '<>', '!=', '>', '>=', '<', '<='].includes(ast.operator)) {
    const first =
      leftNumber === undefined || rightNumber === undefined ? String(leftValue) : leftNumber;
    const second =
      leftNumber === undefined || rightNumber === undefined ? String(rightValue) : rightNumber;
    const value =
      ast.operator === '=' || ast.operator === '=='
        ? first === second
        : ast.operator === '<>' || ast.operator === '!='
          ? first !== second
          : ast.operator === '>'
            ? first > second
            : ast.operator === '>='
              ? first >= second
              : ast.operator === '<'
                ? first < second
                : first <= second;
    return { type: 'boolean', value };
  }
  if (leftNumber === undefined || rightNumber === undefined) {
    return { type: 'error', value: '#VALUE!' };
  }
  if (ast.operator === '/' && rightNumber === 0) return { type: 'error', value: '#DIV/0!' };
  const value =
    ast.operator === '+'
      ? leftNumber + rightNumber
      : ast.operator === '-'
        ? leftNumber - rightNumber
        : ast.operator === '*'
          ? leftNumber * rightNumber
          : leftNumber / rightNumber;
  return Number.isFinite(value) ? { type: 'number', value } : { type: 'error', value: '#NUM!' };
}

function astNodes(ast: FormulaAst): number {
  if (ast.kind === 'binary') return 1 + astNodes(ast.left) + astNodes(ast.right);
  if (ast.kind === 'unary') return 1 + astNodes(ast.operand);
  if (ast.kind === 'call') {
    return 1 + ast.arguments.reduce((count, argument) => count + astNodes(argument), 0);
  }
  return 1;
}

function parseConditionalFormula(source: string, runtime: FormulaRuntime): FormulaAst {
  if (source.length > runtime.maxFormulaLength) {
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      'Conditional formula exceeds its source length limit',
    );
  }
  const cached = runtime.asts.get(source);
  if (cached !== undefined) return cached;
  const ast = parseFormula(source.startsWith('=') ? source : `=${source}`);
  validateFormula(ast);
  if (astNodes(ast) > runtime.maxAstNodes) {
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      'Conditional formula exceeds its AST node limit',
    );
  }
  runtime.asts.set(source, ast);
  return ast;
}

function matches(
  condition: ConditionalExpression,
  input: ConditionalEvaluationInput,
  runtime: FormulaRuntime,
): boolean {
  if (condition.type === 'blank') return input.value.type === 'blank';
  if (condition.type === 'not-blank') return input.value.type !== 'blank';
  if (condition.type === 'text-contains') return input.text.includes(condition.text);
  if (condition.type === 'cell-is') {
    return compare(scalar(input.value), condition.operator, condition.value, condition.value2);
  }
  if (condition.type === 'cell-is-formula') {
    try {
      const first = evaluateFormulaAst(parseConditionalFormula(condition.source, runtime), input, {
        remaining: runtime.maxEvaluationSteps,
      });
      const second =
        condition.source2 === undefined
          ? undefined
          : evaluateFormulaAst(parseConditionalFormula(condition.source2, runtime), input, {
              remaining: runtime.maxEvaluationSteps,
            });
      if (first.type === 'blank' || first.type === 'error') return false;
      if (second?.type === 'blank' || second?.type === 'error') return false;
      return compare(
        scalar(input.value),
        condition.operator,
        first.value,
        second === undefined ? undefined : second.value,
      );
    } catch (error) {
      if (error instanceof ConditionalFormatError) throw error;
      throw new ConditionalFormatError(
        'INVALID_CONDITIONAL_EXPRESSION',
        error instanceof Error ? error.message : String(error),
      );
    }
  }
  try {
    const ast = parseConditionalFormula(condition.source, runtime);
    return truthy(evaluateFormulaAst(ast, input, { remaining: runtime.maxEvaluationSteps }));
  } catch (error) {
    if (error instanceof ConditionalFormatError) throw error;
    throw new ConditionalFormatError(
      'INVALID_CONDITIONAL_EXPRESSION',
      error instanceof Error ? error.message : String(error),
    );
  }
}

function colorChannels(color: string): readonly [number, number, number] | undefined {
  const hex = color.replace(/^#/u, '');
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9a-f]{6}$/iu.test(rgb)) return undefined;
  return [
    Number.parseInt(rgb.slice(0, 2), 16),
    Number.parseInt(rgb.slice(2, 4), 16),
    Number.parseInt(rgb.slice(4, 6), 16),
  ];
}

function interpolateColor(start: string, end: string, ratio: number): string | undefined {
  const first = colorChannels(start);
  const second = colorChannels(end);
  if (first === undefined || second === undefined) return undefined;
  return `#${first
    .map((channel, index) =>
      Math.round(channel + ((second[index] as number) - channel) * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

function colorScalePatch(
  rule: ConditionalRule,
  input: ConditionalEvaluationInput,
  cache: WeakMap<
    ConditionalRule,
    WeakMap<NonNullable<ConditionalEvaluationInput['lookup']>, readonly [number, number]>
  >,
): string | undefined {
  if (rule.effect.type !== 'color-scale') return undefined;
  const current = numeric(
    input.value.type === 'array' ? { type: 'error', value: '#VALUE!' } : input.value,
  );
  if (current === undefined) return undefined;
  let bounds = input.lookup === undefined ? undefined : cache.get(rule)?.get(input.lookup);
  if (bounds === undefined) {
    const values: number[] = [];
    for (const range of rule.ranges) {
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let column = range.start.column; column <= range.end.column; column += 1) {
          const value =
            input.lookup?.({ sheetId: range.sheetId, row, column }) ??
            (row === input.address.row &&
            column === input.address.column &&
            range.sheetId === input.address.sheetId
              ? input.value
              : undefined);
          if (value === undefined || value.type === 'array') continue;
          const number = numeric(value);
          if (number !== undefined) values.push(number);
        }
      }
    }
    if (values.length === 0) return undefined;
    bounds = [Math.min(...values), Math.max(...values)];
    if (input.lookup !== undefined) {
      const byLookup = cache.get(rule) ?? new WeakMap();
      byLookup.set(input.lookup, bounds);
      cache.set(rule, byLookup);
    }
  }
  const [minimum, maximum] = bounds;
  const ratio = maximum === minimum ? 0.5 : (current - minimum) / (maximum - minimum);
  if (rule.effect.midpointColor !== undefined) {
    return ratio <= 0.5
      ? interpolateColor(rule.effect.minimumColor, rule.effect.midpointColor, ratio * 2)
      : interpolateColor(rule.effect.midpointColor, rule.effect.maximumColor, (ratio - 0.5) * 2);
  }
  return interpolateColor(
    rule.effect.minimumColor,
    rule.effect.maximumColor,
    Math.max(0, Math.min(1, ratio)),
  );
}

/** Creates an isolated pure conditional-format evaluator. */
export function createConditionalFormatEvaluator(limits: ConditionalFormatLimits): {
  evaluate(input: ConditionalEvaluationInput): ConditionalEvaluationResult;
} {
  const formulaRuntime: FormulaRuntime = {
    asts: new Map(),
    maxFormulaLength: limits.maxFormulaLength ?? 4_096,
    maxAstNodes: limits.maxAstNodes ?? 1_024,
    maxEvaluationSteps: limits.maxEvaluationSteps ?? 4_096,
  };
  const colorScaleCache = new WeakMap<
    ConditionalRule,
    WeakMap<NonNullable<ConditionalEvaluationInput['lookup']>, readonly [number, number]>
  >();
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
        if (!contains(rule, input.address) || !matches(rule.condition, input, formulaRuntime))
          continue;
        matched.push(rule.id);
        if (rule.effect.type === 'style') Object.assign(patch, rule.effect.patch);
        else {
          const color = colorScalePatch(rule, input, colorScaleCache);
          if (color === undefined)
            diagnostics.push({ code: 'UNSUPPORTED_FORMAT_FEATURE', ruleId: rule.id });
          else patch.backgroundColor = color;
        }
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
