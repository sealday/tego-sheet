/** Supported conditional-section comparison operators. */
export type NumberFormatConditionOperator = '<' | '<=' | '>' | '>=' | '=' | '<>';

/** Numeric predicate that selects a format section. */
export interface NumberFormatCondition {
  /** Comparison operator. */
  readonly operator: NumberFormatConditionOperator;
  /** Finite comparison value. */
  readonly value: number;
}

/** Parsed literal, placeholder, number, or date token. */
export type NumberFormatToken =
  | {
      /** Literal token discriminator. */
      readonly kind: 'literal';
      /** Unescaped literal text. */
      readonly value: string;
    }
  | {
      /** Numeric placeholder discriminator. */
      readonly kind: 'number-pattern';
      /** Excel-oriented numeric placeholder source. */
      readonly value: string;
    }
  | {
      /** Date/time placeholder discriminator. */
      readonly kind: 'date-pattern';
      /** Excel-oriented date/time placeholder source. */
      readonly value: string;
    }
  | {
      /** Text placeholder discriminator. */
      readonly kind: 'text-placeholder';
    };

/** One positive, negative, zero, text, or conditional format section. */
export interface NumberFormatSection {
  /** Optional display color token. */
  readonly color?: string;
  /** Optional numeric selection condition. */
  readonly condition?: NumberFormatCondition;
  /** Parsed display tokens. */
  readonly tokens: readonly NumberFormatToken[];
  /** Original section source. */
  readonly source: string;
}

/** Parsed Excel-oriented custom number format. */
export interface NumberFormatAst {
  /** Original full format code. */
  readonly source: string;
  /** One to four ordered format sections. */
  readonly sections: readonly NumberFormatSection[];
}

/** Syntax error raised for invalid or unsupported format source. */
export class NumberFormatSyntaxError extends SyntaxError {
  /** Stable invalid-format code. */
  readonly code = 'NUMBER_FORMAT_INVALID';
  /** Creates an invalid-format error. */
  constructor(message: string) {
    super(message);
    this.name = 'NumberFormatSyntaxError';
  }
}

const colors = new Set(['Black', 'Blue', 'Cyan', 'Green', 'Magenta', 'Red', 'White', 'Yellow']);

function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  let quoted = false;
  for (let index = 0; index < code.length; index += 1) {
    const character = code[index] as string;
    if (character === '"') quoted = !quoted;
    if (character === ';' && !quoted) {
      sections.push(current);
      current = '';
    } else {
      current += character;
    }
  }
  if (quoted) throw new NumberFormatSyntaxError('Unterminated number-format literal');
  sections.push(current);
  if (sections.length > 4)
    throw new NumberFormatSyntaxError('Number format has more than 4 sections');
  return sections;
}

function parseSection(source: string): NumberFormatSection {
  let remaining = source;
  let color: string | undefined;
  let condition: NumberFormatCondition | undefined;
  while (remaining.startsWith('[')) {
    const close = remaining.indexOf(']');
    if (close < 0) throw new NumberFormatSyntaxError('Unterminated bracket token');
    const content = remaining.slice(1, close);
    const conditionMatch = /^(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)$/u.exec(content);
    if (colors.has(content)) color = content;
    else if (conditionMatch !== null) {
      condition = {
        operator: conditionMatch[1] as NumberFormatConditionOperator,
        value: Number(conditionMatch[2]),
      };
    } else if (/^[hms]+$/iu.test(content)) {
      break;
    } else {
      throw new NumberFormatSyntaxError(`Unsupported bracket token [${content}]`);
    }
    remaining = remaining.slice(close + 1);
  }

  const tokens: NumberFormatToken[] = [];
  let index = 0;
  let buffer = '';
  const flush = (): void => {
    if (buffer.length > 0) {
      tokens.push({ kind: 'literal', value: buffer });
      buffer = '';
    }
  };
  while (index < remaining.length) {
    const character = remaining[index] as string;
    if (character === '"') {
      flush();
      const close = remaining.indexOf('"', index + 1);
      if (close < 0) throw new NumberFormatSyntaxError('Unterminated number-format literal');
      tokens.push({ kind: 'literal', value: remaining.slice(index + 1, close) });
      index = close + 1;
      continue;
    }
    if (character === '\\') {
      const escaped = remaining[index + 1];
      if (escaped === undefined) throw new NumberFormatSyntaxError('Dangling escape');
      buffer += escaped;
      index += 2;
      continue;
    }
    if (character === '@') {
      flush();
      tokens.push({ kind: 'text-placeholder' });
      index += 1;
      continue;
    }
    const elapsed = /^\[h+\]/iu.exec(remaining.slice(index));
    if (elapsed !== null) {
      flush();
      tokens.push({ kind: 'date-pattern', value: elapsed[0] });
      index += elapsed[0].length;
      continue;
    }
    const date = /^(?:yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/iu.exec(
      remaining.slice(index),
    );
    if (date !== null && /[ymdhs]/iu.test(remaining)) {
      flush();
      tokens.push({ kind: 'date-pattern', value: date[0] });
      index += date[0].length;
      continue;
    }
    const number = /^[#0?,.E+%]+/u.exec(remaining.slice(index));
    if (number !== null && /[#0?E%]/iu.test(number[0])) {
      flush();
      tokens.push({ kind: 'number-pattern', value: number[0] });
      index += number[0].length;
      continue;
    }
    buffer += character;
    index += 1;
  }
  flush();
  return {
    ...(color === undefined ? {} : { color }),
    ...(condition === undefined ? {} : { condition }),
    tokens: Object.freeze(tokens),
    source,
  };
}

/** Parses an Excel-oriented number format without locale-dependent guessing. */
export function parseNumberFormat(code: string): NumberFormatAst {
  if (code.length === 0) throw new NumberFormatSyntaxError('Number format is empty');
  const sections = splitSections(code).map(parseSection);
  for (const section of sections) {
    if (section.condition !== undefined) Object.freeze(section.condition);
    for (const token of section.tokens) Object.freeze(token);
    Object.freeze(section);
  }
  return Object.freeze({
    source: code,
    sections: Object.freeze(sections),
  });
}
