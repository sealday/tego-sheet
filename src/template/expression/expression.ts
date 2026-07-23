/** Stable failure raised by the restricted template expression language. */
export class TemplateExpressionError extends SyntaxError {
  readonly code: 'INVALID_EXPRESSION' | 'TEMPLATE_EXPRESSION_UNSAFE' | 'UNKNOWN_FORMATTER';

  constructor(
    code: 'INVALID_EXPRESSION' | 'TEMPLATE_EXPRESSION_UNSAFE' | 'UNKNOWN_FORMATTER',
    message: string,
  ) {
    super(message);
    this.name = 'TemplateExpressionError';
    this.code = code;
  }
}

export type TemplateExpressionNode =
  | { readonly kind: 'literal'; readonly value: null | string | number | boolean }
  | { readonly kind: 'identifier'; readonly name: string }
  | {
      readonly kind: 'member';
      readonly object: TemplateExpressionNode;
      readonly property: string;
    }
  | {
      readonly kind: 'unary';
      readonly operator: '!' | '-';
      readonly operand: TemplateExpressionNode;
    }
  | {
      readonly kind: 'binary';
      readonly operator:
        | '??'
        | '||'
        | '&&'
        | '=='
        | '!='
        | '>'
        | '>='
        | '<'
        | '<='
        | '+'
        | '-'
        | '*'
        | '/';
      readonly left: TemplateExpressionNode;
      readonly right: TemplateExpressionNode;
    }
  | {
      readonly kind: 'conditional';
      readonly test: TemplateExpressionNode;
      readonly consequent: TemplateExpressionNode;
      readonly alternate: TemplateExpressionNode;
    }
  | {
      readonly kind: 'call';
      readonly formatter: string;
      readonly arguments: readonly TemplateExpressionNode[];
    };

export interface CompiledTemplateExpression {
  readonly source: string;
  readonly ast: TemplateExpressionNode;
}

export interface TemplateExpressionScope {
  readonly root: unknown;
  readonly item?: unknown;
  readonly index?: number;
  readonly first?: boolean;
  readonly last?: boolean;
}

export type TemplateFormatter = (...values: readonly unknown[]) => unknown;
export type TemplateFormatterRegistry = Readonly<Record<string, TemplateFormatter>>;

interface Token {
  readonly kind: 'number' | 'string' | 'identifier' | 'operator' | 'punctuation' | 'eof';
  readonly value: string;
  readonly offset: number;
}

const BLOCKED_PROPERTIES = new Set(['__proto__', 'prototype', 'constructor']);
const BLOCKED_IDENTIFIERS = new Set([
  'globalThis',
  'window',
  'document',
  'Function',
  'eval',
  'process',
  'require',
  'import',
  'this',
  'new',
]);
const OPERATORS = ['??', '||', '&&', '==', '!=', '>=', '<=', '>', '<', '+', '-', '*', '/', '!'];
const PRECEDENCE: Readonly<Record<string, number | undefined>> = {
  '??': 1,
  '||': 2,
  '&&': 3,
  '==': 4,
  '!=': 4,
  '>': 5,
  '>=': 5,
  '<': 5,
  '<=': 5,
  '+': 6,
  '-': 6,
  '*': 7,
  '/': 7,
};

function freezeNode<T extends TemplateExpressionNode>(node: T): T {
  if (node.kind === 'member') freezeNode(node.object);
  if (node.kind === 'unary') freezeNode(node.operand);
  if (node.kind === 'binary') {
    freezeNode(node.left);
    freezeNode(node.right);
  }
  if (node.kind === 'conditional') {
    freezeNode(node.test);
    freezeNode(node.consequent);
    freezeNode(node.alternate);
  }
  if (node.kind === 'call') {
    node.arguments.forEach(freezeNode);
    Object.freeze(node.arguments);
  }
  return Object.freeze(node);
}

function tokenize(source: string): readonly Token[] {
  const tokens: Token[] = [];
  let offset = 0;
  while (offset < source.length) {
    const character = source[offset]!;
    if (/\s/u.test(character)) {
      offset += 1;
      continue;
    }
    if (character === '"' || character === "'") {
      const quote = character;
      const start = offset++;
      let value = '';
      let closed = false;
      while (offset < source.length) {
        const next = source[offset++]!;
        if (next === quote) {
          closed = true;
          break;
        }
        if (next === '\\') {
          const escaped = source[offset++];
          if (escaped === undefined) break;
          value += escaped === 'n' ? '\n' : escaped === 't' ? '\t' : escaped;
        } else {
          value += next;
        }
      }
      if (!closed) throw new TemplateExpressionError('INVALID_EXPRESSION', 'Unterminated string');
      tokens.push({ kind: 'string', value, offset: start });
      continue;
    }
    const number = source.slice(offset).match(/^(?:\d+(?:\.\d+)?|\.\d+)/u)?.[0];
    if (number !== undefined) {
      tokens.push({ kind: 'number', value: number, offset });
      offset += number.length;
      continue;
    }
    const identifier = source.slice(offset).match(/^(?:[$_\p{L}][$_\p{L}\p{N}]*)/u)?.[0];
    if (identifier !== undefined) {
      if (BLOCKED_IDENTIFIERS.has(identifier)) {
        throw new TemplateExpressionError(
          'TEMPLATE_EXPRESSION_UNSAFE',
          `Identifier ${identifier} is not available`,
        );
      }
      tokens.push({ kind: 'identifier', value: identifier, offset });
      offset += identifier.length;
      continue;
    }
    const operator = OPERATORS.find((candidate) => source.startsWith(candidate, offset));
    if (operator !== undefined) {
      tokens.push({ kind: 'operator', value: operator, offset });
      offset += operator.length;
      continue;
    }
    if ('().,?:'.includes(character)) {
      tokens.push({ kind: 'punctuation', value: character, offset });
      offset += 1;
      continue;
    }
    if (character === '=' || character === '[' || character === ']' || character === '{') {
      throw new TemplateExpressionError(
        'TEMPLATE_EXPRESSION_UNSAFE',
        `Unsafe expression token at ${offset}`,
      );
    }
    throw new TemplateExpressionError(
      'INVALID_EXPRESSION',
      `Unexpected expression token at ${offset}`,
    );
  }
  tokens.push({ kind: 'eof', value: '', offset });
  return tokens;
}

class Parser {
  readonly #tokens: readonly Token[];
  #index = 0;

  constructor(tokens: readonly Token[]) {
    this.#tokens = tokens;
  }

  parse(): TemplateExpressionNode {
    const expression = this.#conditional();
    if (this.#current().kind !== 'eof') this.#invalid('Unexpected trailing expression input');
    return freezeNode(expression);
  }

  #current(): Token {
    return this.#tokens[this.#index]!;
  }

  #consume(value?: string): Token {
    const token = this.#current();
    if (value !== undefined && token.value !== value) this.#invalid(`Expected ${value}`);
    this.#index += 1;
    return token;
  }

  #invalid(message: string): never {
    throw new TemplateExpressionError(
      'INVALID_EXPRESSION',
      `${message} at ${this.#current().offset}`,
    );
  }

  #conditional(): TemplateExpressionNode {
    const test = this.#binary(1);
    if (this.#current().value !== '?') return test;
    this.#consume('?');
    const consequent = this.#conditional();
    this.#consume(':');
    return {
      kind: 'conditional',
      test,
      consequent,
      alternate: this.#conditional(),
    };
  }

  #binary(minimum: number): TemplateExpressionNode {
    let left = this.#unary();
    while (true) {
      const token = this.#current();
      const precedence = PRECEDENCE[token.value];
      if (precedence === undefined || precedence < minimum) return left;
      this.#consume();
      left = {
        kind: 'binary',
        operator: token.value as Extract<TemplateExpressionNode, { kind: 'binary' }>['operator'],
        left,
        right: this.#binary(precedence + 1),
      };
    }
  }

  #unary(): TemplateExpressionNode {
    const token = this.#current();
    if (token.value === '!' || token.value === '-') {
      this.#consume();
      return { kind: 'unary', operator: token.value, operand: this.#unary() };
    }
    return this.#postfix();
  }

  #postfix(): TemplateExpressionNode {
    let expression = this.#primary();
    while (true) {
      if (this.#current().value === '.') {
        this.#consume('.');
        const property = this.#consume();
        if (property.kind !== 'identifier') this.#invalid('Expected property name');
        if (BLOCKED_PROPERTIES.has(property.value)) {
          throw new TemplateExpressionError(
            'TEMPLATE_EXPRESSION_UNSAFE',
            `Property ${property.value} is not available`,
          );
        }
        expression = { kind: 'member', object: expression, property: property.value };
        continue;
      }
      if (this.#current().value === '(') {
        if (expression.kind !== 'identifier') {
          throw new TemplateExpressionError(
            'TEMPLATE_EXPRESSION_UNSAFE',
            'Only registered formatter calls are allowed',
          );
        }
        this.#consume('(');
        const args: TemplateExpressionNode[] = [];
        if (this.#current().value !== ')') {
          while (this.#current().value !== ')') {
            args.push(this.#conditional());
            if (this.#current().value !== ',') break;
            this.#consume(',');
          }
        }
        this.#consume(')');
        expression = { kind: 'call', formatter: expression.name, arguments: args };
        continue;
      }
      return expression;
    }
  }

  #primary(): TemplateExpressionNode {
    const token = this.#consume();
    if (token.kind === 'number') return { kind: 'literal', value: Number(token.value) };
    if (token.kind === 'string') return { kind: 'literal', value: token.value };
    if (token.kind === 'identifier') {
      if (token.value === 'true') return { kind: 'literal', value: true };
      if (token.value === 'false') return { kind: 'literal', value: false };
      if (token.value === 'null') return { kind: 'literal', value: null };
      return { kind: 'identifier', name: token.value };
    }
    if (token.value === '(') {
      const expression = this.#conditional();
      this.#consume(')');
      return expression;
    }
    this.#invalid('Expected an expression');
  }
}

/** Parses a source string into an immutable, non-executable expression IR. */
export function compileTemplateExpression(source: string): CompiledTemplateExpression {
  if (typeof source !== 'string' || source.trim() === '') {
    throw new TemplateExpressionError('INVALID_EXPRESSION', 'Expression source is empty');
  }
  return Object.freeze({ source, ast: new Parser(tokenize(source)).parse() });
}

function frozenValue(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
  if (value === null || typeof value !== 'object') return value;
  const existing = seen.get(value);
  if (existing !== undefined) return existing;
  const output: unknown[] | Record<string, unknown> = Array.isArray(value) ? [] : {};
  seen.set(value, output);
  if (Array.isArray(value)) {
    value.forEach((item) => (output as unknown[]).push(frozenValue(item, seen)));
  } else {
    for (const [key, item] of Object.entries(value)) {
      if (BLOCKED_PROPERTIES.has(key)) continue;
      Object.defineProperty(output, key, {
        enumerable: true,
        value: frozenValue(item, seen),
      });
    }
  }
  return Object.freeze(output);
}

function readMember(value: unknown, property: string): unknown {
  if (BLOCKED_PROPERTIES.has(property) || value === null || typeof value !== 'object') {
    return undefined;
  }
  if (!Object.prototype.hasOwnProperty.call(value, property)) return undefined;
  return (value as Readonly<Record<string, unknown>>)[property];
}

function evaluate(
  node: TemplateExpressionNode,
  variables: Readonly<Record<string, unknown>>,
  formatters: TemplateFormatterRegistry,
): unknown {
  if (node.kind === 'literal') return node.value;
  if (node.kind === 'identifier') return variables[node.name];
  if (node.kind === 'member')
    return readMember(evaluate(node.object, variables, formatters), node.property);
  if (node.kind === 'unary') {
    const value = evaluate(node.operand, variables, formatters);
    return node.operator === '!' ? !value : -Number(value);
  }
  if (node.kind === 'conditional') {
    return evaluate(node.test, variables, formatters)
      ? evaluate(node.consequent, variables, formatters)
      : evaluate(node.alternate, variables, formatters);
  }
  if (node.kind === 'call') {
    const formatter = formatters[node.formatter];
    if (formatter === undefined) {
      throw new TemplateExpressionError(
        'UNKNOWN_FORMATTER',
        `Unknown formatter: ${node.formatter}`,
      );
    }
    return formatter(
      ...node.arguments.map((argument) => frozenValue(evaluate(argument, variables, formatters))),
    );
  }
  const left = evaluate(node.left, variables, formatters);
  if (node.operator === '??') return left ?? evaluate(node.right, variables, formatters);
  if (node.operator === '||') return left || evaluate(node.right, variables, formatters);
  if (node.operator === '&&') return left && evaluate(node.right, variables, formatters);
  const right = evaluate(node.right, variables, formatters);
  switch (node.operator) {
    case '==':
      return left === right;
    case '!=':
      return left !== right;
    case '>':
      return (left as number) > (right as number);
    case '>=':
      return (left as number) >= (right as number);
    case '<':
      return (left as number) < (right as number);
    case '<=':
      return (left as number) <= (right as number);
    case '+':
      return typeof left === 'string' || typeof right === 'string'
        ? `${String(left)}${String(right)}`
        : Number(left) + Number(right);
    case '-':
      return Number(left) - Number(right);
    case '*':
      return Number(left) * Number(right);
    case '/':
      return Number(left) / Number(right);
  }
}

/** Interprets validated IR against an immutable, capability-limited scope. */
export function evaluateTemplateExpression(
  expression: CompiledTemplateExpression,
  scope: TemplateExpressionScope,
  formatters: TemplateFormatterRegistry = {},
): unknown {
  const root = frozenValue(scope.root);
  const variables = Object.freeze({
    ...(root !== null && typeof root === 'object' && !Array.isArray(root) ? root : {}),
    root,
    item: frozenValue(scope.item),
    $index: scope.index ?? 0,
    $first: scope.first ?? false,
    $last: scope.last ?? false,
  });
  return frozenValue(evaluate(expression.ast, variables, Object.freeze({ ...formatters })));
}
