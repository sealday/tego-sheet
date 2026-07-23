import type { FormulaAst, FormulaReference, SourceSpan } from './ast';
import { FormulaSyntaxError, freezeFormulaAst } from './ast';

type TokenKind =
  | 'number'
  | 'string'
  | 'word'
  | 'reference'
  | 'error'
  | 'operator'
  | 'left'
  | 'right'
  | 'comma'
  | 'colon'
  | 'bang'
  | 'eof';

interface Token {
  readonly kind: TokenKind;
  readonly value: string;
  readonly span: SourceSpan;
}

const referencePattern = /^\$?[A-Z]+\$?[1-9]\d*/i;
const wordPattern = /^[A-Z_][A-Z0-9_.]*/i;
const numberPattern = /^(?:\d+(?:\.\d*)?|\.\d+)/;

function tokenize(source: string): readonly Token[] {
  const tokens: Token[] = [];
  let index = source.startsWith('=') ? 1 : 0;
  while (index < source.length) {
    const start = index;
    const character = source[index] as string;
    if (/\s/u.test(character)) {
      index += 1;
      continue;
    }
    if (character === '#') {
      const error = /^(?:#REF!|#VALUE!|#DIV\/0!|#NAME\?|#N\/A|#NUM!|#SPILL!)/u.exec(
        source.slice(index),
      );
      if (error === null)
        throw new FormulaSyntaxError('Unknown formula error literal', {
          start,
          end: start + 1,
        });
      index += error[0].length;
      tokens.push({ kind: 'error', value: error[0], span: { start, end: index } });
      continue;
    }
    if (character === '"' || character === "'") {
      const quote = character;
      let value = '';
      index += 1;
      while (index < source.length) {
        if (source[index] === quote) {
          if (source[index + 1] === quote) {
            value += quote;
            index += 2;
            continue;
          }
          break;
        }
        value += source[index];
        index += 1;
      }
      if (source[index] !== quote)
        throw new FormulaSyntaxError('Unterminated quoted value', { start, end: source.length });
      index += 1;
      tokens.push({
        kind: quote === '"' ? 'string' : 'word',
        value,
        span: { start, end: index },
      });
      continue;
    }
    const number = numberPattern.exec(source.slice(index));
    if (number !== null) {
      index += number[0].length;
      tokens.push({ kind: 'number', value: number[0], span: { start, end: index } });
      continue;
    }
    const possibleFunction = wordPattern.exec(source.slice(index));
    if (
      possibleFunction !== null &&
      /^\s*\(/u.test(source.slice(index + possibleFunction[0].length))
    ) {
      index += possibleFunction[0].length;
      tokens.push({
        kind: 'word',
        value: possibleFunction[0],
        span: { start, end: index },
      });
      continue;
    }
    const reference = referencePattern.exec(source.slice(index));
    if (reference !== null) {
      const following = source[index + reference[0].length];
      if (following === '!') {
        index += reference[0].length;
        tokens.push({
          kind: 'word',
          value: reference[0],
          span: { start, end: index },
        });
        continue;
      }
      if (following === undefined || !/[A-Z0-9_]/iu.test(following)) {
        index += reference[0].length;
        tokens.push({
          kind: 'reference',
          value: reference[0].toUpperCase(),
          span: { start, end: index },
        });
        continue;
      }
    }
    const word = wordPattern.exec(source.slice(index));
    if (word !== null) {
      index += word[0].length;
      tokens.push({ kind: 'word', value: word[0], span: { start, end: index } });
      continue;
    }
    const two = source.slice(index, index + 2);
    if (['>=', '<=', '<>', '!=', '=='].includes(two)) {
      index += 2;
      tokens.push({ kind: 'operator', value: two, span: { start, end: index } });
      continue;
    }
    const punctuation: Readonly<Record<string, TokenKind | undefined>> = {
      '(': 'left',
      ')': 'right',
      ',': 'comma',
      ':': 'colon',
      '!': 'bang',
    };
    const punctuationKind = punctuation[character];
    if (punctuationKind !== undefined) {
      index += 1;
      tokens.push({ kind: punctuationKind, value: character, span: { start, end: index } });
      continue;
    }
    if ('+-*/=><&'.includes(character)) {
      index += 1;
      tokens.push({ kind: 'operator', value: character, span: { start, end: index } });
      continue;
    }
    throw new FormulaSyntaxError(`Unexpected character ${character}`, { start, end: start + 1 });
  }
  tokens.push({ kind: 'eof', value: '', span: { start: source.length, end: source.length } });
  return tokens;
}

function columnIndex(label: string): number {
  let result = 0;
  for (const character of label) result = result * 26 + character.charCodeAt(0) - 64;
  return result - 1;
}

function referenceFrom(token: Token, sheetToken?: string): FormulaReference {
  const match = /^(\$?)([A-Z]+)(\$?)([1-9]\d*)$/u.exec(token.value);
  if (match === null) throw new FormulaSyntaxError('Invalid A1 reference', token.span);
  const column = columnIndex(match[2] as string);
  const row = Number(match[4]) - 1;
  if (
    !Number.isSafeInteger(column) ||
    !Number.isSafeInteger(row) ||
    column < 0 ||
    column >= 16_384 ||
    row < 0 ||
    row >= 1_048_576
  ) {
    throw new FormulaSyntaxError('Reference exceeds Excel worksheet coordinates', token.span);
  }
  return {
    ...(sheetToken === undefined ? {} : { sheetToken }),
    columnAbsolute: match[1] === '$',
    column,
    rowAbsolute: match[3] === '$',
    row,
  };
}

const precedence: Readonly<Record<string, number | undefined>> = {
  '=': 1,
  '==': 1,
  '<>': 1,
  '!=': 1,
  '>': 1,
  '>=': 1,
  '<': 1,
  '<=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
};

class Parser {
  #index = 0;
  constructor(private readonly tokens: readonly Token[]) {}

  parse(): FormulaAst {
    const result = this.binary(1);
    this.consume('eof');
    return result;
  }

  private current(): Token {
    return this.tokens[this.#index] as Token;
  }

  private consume(kind: TokenKind): Token {
    const token = this.current();
    if (token.kind !== kind)
      throw new FormulaSyntaxError(`Expected ${kind}, received ${token.kind}`, token.span);
    this.#index += 1;
    return token;
  }

  private binary(minimum: number): FormulaAst {
    let left = this.unary();
    while (this.current().kind === 'operator') {
      const token = this.current();
      const level = precedence[token.value] ?? 0;
      if (level < minimum) break;
      this.#index += 1;
      const right = this.binary(level + 1);
      left = {
        kind: 'binary',
        operator: token.value as Extract<FormulaAst, { kind: 'binary' }>['operator'],
        left,
        right,
        span: { start: left.span.start, end: right.span.end },
      };
    }
    return left;
  }

  private unary(): FormulaAst {
    const token = this.current();
    if (token.kind === 'operator' && token.value === '-') {
      this.#index += 1;
      const operand = this.unary();
      return {
        kind: 'unary',
        operator: '-',
        operand,
        span: { start: token.span.start, end: operand.span.end },
      };
    }
    return this.primary();
  }

  private primary(): FormulaAst {
    const token = this.current();
    if (token.kind === 'number') {
      this.#index += 1;
      return { kind: 'number', value: Number(token.value), span: token.span };
    }
    if (token.kind === 'string') {
      this.#index += 1;
      return { kind: 'string', value: token.value, span: token.span };
    }
    if (token.kind === 'error') {
      this.#index += 1;
      return {
        kind: 'error',
        value: token.value as Extract<FormulaAst, { kind: 'error' }>['value'],
        span: token.span,
      };
    }
    if (token.kind === 'left') {
      this.#index += 1;
      const value = this.binary(1);
      const end = this.consume('right');
      return { ...value, span: { start: token.span.start, end: end.span.end } };
    }
    if (token.kind === 'reference') return this.reference();
    if (token.kind === 'word') {
      this.#index += 1;
      if (this.current().kind === 'bang') {
        this.#index += 1;
        return this.reference(token);
      }
      if (token.value === 'TRUE' || token.value === 'FALSE') {
        return { kind: 'boolean', value: token.value === 'TRUE', span: token.span };
      }
      if (this.current().kind !== 'left')
        throw new FormulaSyntaxError(`Unknown identifier ${token.value}`, token.span);
      this.#index += 1;
      const arguments_: FormulaAst[] = [];
      if (this.current().kind !== 'right') {
        arguments_.push(this.binary(1));
        while (this.current().kind === 'comma') {
          this.#index += 1;
          arguments_.push(this.binary(1));
        }
      }
      const end = this.consume('right');
      return {
        kind: 'call',
        name: token.value.toUpperCase(),
        arguments: arguments_,
        span: { start: token.span.start, end: end.span.end },
      };
    }
    throw new FormulaSyntaxError('Expected expression', token.span);
  }

  private reference(sheet?: Token): FormulaAst {
    const startToken = this.consume('reference');
    const start = referenceFrom(startToken, sheet?.value);
    if (this.current().kind !== 'colon') {
      return {
        kind: 'reference',
        reference: start,
        span: { start: sheet?.span.start ?? startToken.span.start, end: startToken.span.end },
      };
    }
    this.#index += 1;
    let endSheet: Token | undefined;
    if (this.current().kind === 'word') {
      endSheet = this.current();
      this.#index += 1;
      this.consume('bang');
    }
    const endToken = this.consume('reference');
    return {
      kind: 'range',
      start,
      end: referenceFrom(endToken, endSheet?.value ?? sheet?.value),
      span: { start: sheet?.span.start ?? startToken.span.start, end: endToken.span.end },
    };
  }
}

/** Parses restricted formula source into a typed, non-executable AST. */
export function parseFormula(source: string): FormulaAst {
  if (source.length > 8192) {
    throw new FormulaSyntaxError('Formula exceeds the 8192 character limit', {
      start: 8192,
      end: source.length,
    });
  }
  if (!source.startsWith('=')) {
    throw new FormulaSyntaxError('Formula source must start with =', { start: 0, end: 0 });
  }
  const tokens = tokenize(source);
  if (tokens.length > 4096) {
    throw new FormulaSyntaxError('Formula exceeds the 4096 token limit', {
      start: 0,
      end: source.length,
    });
  }
  let depth = 0;
  for (const token of tokens) {
    if (token.kind === 'left') {
      depth += 1;
      if (depth > 256) {
        throw new FormulaSyntaxError('Formula exceeds the nesting limit', token.span);
      }
    } else if (token.kind === 'right') {
      depth -= 1;
    }
  }
  return freezeFormulaAst(new Parser(tokens).parse());
}
