import { describe, expect, it } from 'vitest';
import {
  BUILTIN_NUMBER_FORMAT_COMPATIBILITY,
  createNumberFormatter,
} from '../../../src/format/number-formatter';
import { parseNumberFormat } from '../../../src/format/number-format-parser';

const formatter = createNumberFormatter();
const context = {
  locale: 'en-US',
  timeZone: 'UTC',
  dateSystem: 'excel-1900' as const,
};

describe('Excel-oriented number formats', () => {
  it('parses four sections, conditions, colors, escaped literals, and date tokens', () => {
    const ast = parseNumberFormat('[Red][<0]$#,##0.00;[Blue]0%;0.00;"text:" @');
    expect(ast.sections).toHaveLength(4);
    expect(ast.sections[0]).toMatchObject({ color: 'Red', condition: { operator: '<', value: 0 } });
    expect(ast.sections[3]?.tokens).toContainEqual({ kind: 'text-placeholder' });
    expect(Object.isFrozen(ast)).toBe(true);
    expect(Object.isFrozen(ast.sections[0]?.tokens[0])).toBe(true);
  });

  it('formats number, currency, percentage, scientific, date, time, and text deterministically', () => {
    expect(formatter.format({ type: 'number', value: 1234.5 }, '#,##0.00', context)).toBe(
      '1,234.50',
    );
    expect(formatter.format({ type: 'number', value: -12.5 }, '$#,##0.00', context)).toBe(
      '-$12.50',
    );
    expect(formatter.format({ type: 'number', value: 0.125 }, '0.0%', context)).toBe('12.5%');
    expect(formatter.format({ type: 'number', value: 12345 }, '0.00E+00', context)).toBe(
      '1.23E+04',
    );
    expect(formatter.format({ type: 'number', value: 45293 }, 'yyyy-mm-dd', context)).toBe(
      '2024-01-02',
    );
    expect(formatter.format({ type: 'number', value: 0.5 }, 'hh:mm:ss', context)).toBe('12:00:00');
    expect(formatter.format({ type: 'string', value: 'ok' }, '0;0;0;"value=" @', context)).toBe(
      'value= ok',
    );
  });

  it('honors date systems and explicit locale/time zone without consulting host defaults', () => {
    const serial = { type: 'number' as const, value: 1 };
    expect(formatter.format(serial, 'yyyy-mm-dd', context)).toBe('1900-01-01');
    expect(formatter.format(serial, 'yyyy-mm-dd', { ...context, dateSystem: 'excel-1904' })).toBe(
      '1904-01-02',
    );
    expect(
      formatter.format({ type: 'number', value: 45293.5 }, 'yyyy-mm-dd hh:mm', {
        ...context,
        locale: 'zh-CN',
        timeZone: 'Asia/Shanghai',
      }),
    ).toBe('2024-01-02 12:00');
    expect(formatter.format({ type: 'number', value: 60 }, 'yyyy-mm-dd', context)).toBe(
      '1900-02-29',
    );
  });

  it('rejects unsupported tokens instead of guessing and declares compatibility', () => {
    expect(() => parseNumberFormat('[DBNum1]0')).toThrow(
      expect.objectContaining({ code: 'NUMBER_FORMAT_INVALID' }),
    );
    expect(BUILTIN_NUMBER_FORMAT_COMPATIBILITY).toContainEqual({
      code: 'yyyy-mm-dd',
      category: 'date',
    });
  });
});
