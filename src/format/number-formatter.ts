import type { FormulaValue } from '../formula/ast';
import type {
  NumberFormatAst,
  NumberFormatCondition,
  NumberFormatSection,
  NumberFormatToken,
} from './number-format-parser';
import { parseNumberFormat } from './number-format-parser';

/** Explicit deterministic inputs for number and date formatting. */
export interface FormatContext {
  /** BCP 47 locale used for separators and grouping. */
  readonly locale: string;
  /** IANA time-zone identifier supplied by the host. */
  readonly timeZone: string;
  /** Workbook Excel serial-date system. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
}

/** Parser and deterministic renderer for Excel-oriented number formats. */
export interface NumberFormatter {
  /** Parses and caches a custom number format. */
  parse(code: string): NumberFormatAst;
  /** Formats a typed formula value with explicit context. */
  format(value: FormulaValue, code: string, context: FormatContext): string;
}

/** One number-format fixture promised compatible by the core. */
export interface NumberFormatCompatibility {
  /** Canonical format code. */
  readonly code: string;
  /** Supported format category. */
  readonly category: 'number' | 'currency' | 'percentage' | 'scientific' | 'date' | 'time' | 'text';
}

/** Stable compatibility manifest for built-in number-format fixtures. */
export const BUILTIN_NUMBER_FORMAT_COMPATIBILITY: readonly NumberFormatCompatibility[] =
  Object.freeze([
    { code: '#,##0.00', category: 'number' },
    { code: '$#,##0.00', category: 'currency' },
    { code: '0.0%', category: 'percentage' },
    { code: '0.00E+00', category: 'scientific' },
    { code: 'yyyy-mm-dd', category: 'date' },
    { code: 'hh:mm:ss', category: 'time' },
    { code: '@', category: 'text' },
  ]);

function matches(condition: NumberFormatCondition, value: number): boolean {
  if (condition.operator === '<') return value < condition.value;
  if (condition.operator === '<=') return value <= condition.value;
  if (condition.operator === '>') return value > condition.value;
  if (condition.operator === '>=') return value >= condition.value;
  if (condition.operator === '=') return value === condition.value;
  return value !== condition.value;
}

function selectSection(
  ast: NumberFormatAst,
  value: FormulaValue,
): { readonly section: NumberFormatSection; readonly useAbsolute: boolean } {
  const conditioned = ast.sections.find(
    (section) =>
      section.condition !== undefined &&
      value.type === 'number' &&
      matches(section.condition, value.value),
  );
  if (conditioned !== undefined) return { section: conditioned, useAbsolute: false };
  if (value.type === 'string' || value.type === 'boolean' || value.type === 'blank') {
    return { section: ast.sections[3] ?? ast.sections[0]!, useAbsolute: false };
  }
  if (value.type !== 'number') return { section: ast.sections[0]!, useAbsolute: false };
  if (value.value > 0) return { section: ast.sections[0]!, useAbsolute: false };
  if (value.value < 0)
    return {
      section: ast.sections[1] ?? ast.sections[0]!,
      useAbsolute: ast.sections[1] !== undefined,
    };
  return { section: ast.sections[2] ?? ast.sections[0]!, useAbsolute: false };
}

function serialParts(
  serial: number,
  dateSystem: 'excel-1900' | 'excel-1904',
): {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
} {
  const whole = Math.floor(serial);
  if (dateSystem === 'excel-1900' && whole === 60) {
    let fakeRemaining = Math.round((serial - whole) * 86_400);
    const fakeHour = Math.floor(fakeRemaining / 3600);
    fakeRemaining -= fakeHour * 3600;
    const fakeMinute = Math.floor(fakeRemaining / 60);
    return {
      year: 1900,
      month: 2,
      day: 29,
      hour: fakeHour,
      minute: fakeMinute,
      second: fakeRemaining - fakeMinute * 60,
    };
  }
  const adjusted = dateSystem === 'excel-1900' && whole >= 60 ? whole - 1 : whole;
  const epoch = dateSystem === 'excel-1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  const date = new Date(epoch + adjusted * 86_400_000);
  let remaining = Math.round((serial - whole) * 86_400);
  const hour = Math.floor(remaining / 3600);
  remaining -= hour * 3600;
  const minute = Math.floor(remaining / 60);
  const second = remaining - minute * 60;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour,
    minute,
    second,
  };
}

function pad(value: number, length = 2): string {
  return String(value).padStart(length, '0');
}

function formatDateToken(
  token: string,
  parts: ReturnType<typeof serialParts>,
  isMinute: boolean,
  serial: number,
  locale: string,
): string {
  const lower = token.toLowerCase();
  if (/^\[h+\]$/u.test(lower)) return String(Math.floor(serial * 24));
  if (lower === 'yyyy') return pad(parts.year, 4);
  if (lower === 'yy') return pad(parts.year % 100);
  if (lower === 'ddd' || lower === 'dddd') {
    return new Intl.DateTimeFormat(locale, {
      weekday: lower === 'ddd' ? 'short' : 'long',
      timeZone: 'UTC',
    }).format(Date.UTC(parts.year, parts.month - 1, parts.day));
  }
  if (lower.startsWith('d')) return lower === 'dd' ? pad(parts.day) : String(parts.day);
  if (lower.startsWith('h')) return lower === 'hh' ? pad(parts.hour) : String(parts.hour);
  if (lower.startsWith('s')) return lower === 'ss' ? pad(parts.second) : String(parts.second);
  if (!isMinute && (lower === 'mmm' || lower === 'mmmm')) {
    return new Intl.DateTimeFormat(locale, {
      month: lower === 'mmm' ? 'short' : 'long',
      timeZone: 'UTC',
    }).format(Date.UTC(parts.year, parts.month - 1, parts.day));
  }
  const value = isMinute ? parts.minute : parts.month;
  return lower === 'mm' ? pad(value) : String(value);
}

function decimalPlaces(pattern: string): number {
  const mantissa = pattern.split(/[Ee]/u)[0] as string;
  const decimal = mantissa.indexOf('.');
  if (decimal < 0) return 0;
  return mantissa.slice(decimal + 1).match(/0/gu)?.length ?? 0;
}

function renderNumber(pattern: string, value: number, locale: string): string {
  const percent = pattern.includes('%');
  const scientific = /E[+-]0+/iu.test(pattern);
  const scaled = percent ? value * 100 : value;
  const digits = decimalPlaces(pattern);
  if (scientific) {
    const [mantissa, exponent = '0'] = scaled.toExponential(digits).split('e');
    const sign = Number(exponent) >= 0 ? '+' : '-';
    const width = /E[+-](0+)/iu.exec(pattern)?.[1].length ?? 2;
    return `${mantissa}E${sign}${pad(Math.abs(Number(exponent)), width)}${percent ? '%' : ''}`;
  }
  const grouping = pattern.includes(',');
  const formatted = new Intl.NumberFormat(locale, {
    useGrouping: grouping,
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(scaled);
  return `${formatted}${percent ? '%' : ''}`;
}

function plain(value: FormulaValue): string {
  if (value.type === 'blank') return '';
  if (value.type === 'array') return '#SPILL!';
  return String(value.value);
}

function formatValue(value: FormulaValue, ast: NumberFormatAst, context: FormatContext): string {
  if (value.type === 'error' || value.type === 'array') return plain(value);
  const { section, useAbsolute } = selectSection(ast, value);
  const sourceNumber =
    value.type === 'number' ? (useAbsolute ? Math.abs(value.value) : value.value) : 0;
  const dateMode = section.tokens.some(({ kind }) => kind === 'date-pattern');
  const parts = dateMode ? serialParts(sourceNumber, context.dateSystem) : undefined;
  let output = '';
  let renderedNumber = false;
  for (const [tokenIndex, token] of section.tokens.entries()) {
    if (token.kind === 'literal') {
      output += token.value;
      continue;
    }
    if (token.kind === 'text-placeholder') {
      output += plain(value);
      continue;
    }
    if (token.kind === 'date-pattern') {
      const previousDate = section.tokens
        .slice(0, tokenIndex)
        .reverse()
        .find(({ kind }) => kind === 'date-pattern');
      const nextDate = section.tokens
        .slice(tokenIndex + 1)
        .find(({ kind }) => kind === 'date-pattern');
      const minuteContext =
        /^m{1,2}$/iu.test(token.value) &&
        ((previousDate?.kind === 'date-pattern' && /^\[?h/iu.test(previousDate.value)) ||
          (nextDate?.kind === 'date-pattern' && /^s/iu.test(nextDate.value)));
      output += formatDateToken(
        token.value,
        parts as ReturnType<typeof serialParts>,
        minuteContext,
        sourceNumber,
        context.locale,
      );
      continue;
    }
    if (!renderedNumber && value.type === 'number') {
      output += renderNumber(token.value, sourceNumber, context.locale);
      renderedNumber = true;
    }
  }
  if (!dateMode && value.type === 'number' && !renderedNumber) output += String(sourceNumber);
  if (
    value.type === 'number' &&
    value.value < 0 &&
    !useAbsolute &&
    !output.startsWith('-') &&
    ast.sections.length === 1
  ) {
    output = `-${output.replace('-', '')}`;
  }
  return output;
}

/** Creates an isolated number formatter with a local parsed-format cache. */
export function createNumberFormatter(): NumberFormatter {
  const cache = new Map<string, NumberFormatAst>();
  return {
    parse(code) {
      const cached = cache.get(code);
      if (cached !== undefined) return cached;
      const ast = parseNumberFormat(code);
      cache.set(code, ast);
      return ast;
    },
    format(value, code, context) {
      return formatValue(value, this.parse(code), context);
    },
  };
}

/** Serializes one parsed token for diagnostics and tooling. */
export function renderNumberFormatToken(token: NumberFormatToken): string {
  if (token.kind === 'text-placeholder') return '@';
  return token.value;
}
