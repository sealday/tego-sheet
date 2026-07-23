import { describe, expect, it } from 'vitest';
import {
  compileTemplateExpression,
  evaluateTemplateExpression,
  TemplateExpressionError,
} from '../../../src/template/expression';

describe('template expression DSL', () => {
  it('compiles and evaluates deterministic data expressions', () => {
    const expression = compileTemplateExpression(
      'customer.name ?? "Anonymous" + " #" + (order.total > 10 ? order.id : "small")',
    );

    expect(
      evaluateTemplateExpression(expression, {
        root: {
          customer: { name: null },
          order: { id: 'A-7', total: 12 },
        },
      }),
    ).toBe('Anonymous #A-7');
  });

  it.each([
    'value.constructor.constructor("return globalThis")()',
    'globalThis.document',
    'value.__proto__',
    'value["constructor"]',
    'value = 1',
    'Function("return 1")()',
  ])('rejects unsafe source %s', (source) => {
    expect(() => compileTemplateExpression(source)).toThrowError(
      expect.objectContaining({ code: 'TEMPLATE_EXPRESSION_UNSAFE' }),
    );
  });

  it('only invokes explicitly registered pure formatters with frozen inputs', () => {
    const expression = compileTemplateExpression('money(order.total)');
    const result = evaluateTemplateExpression(
      expression,
      { root: { order: { total: 12.5 } } },
      {
        money(value) {
          expect(Object.isFrozen(value)).toBe(true);
          return `$${String(value)}`;
        },
      },
    );
    expect(result).toBe('$12.5');
    expect(() => evaluateTemplateExpression(expression, { root: { order: { total: 1 } } })).toThrow(
      TemplateExpressionError,
    );
  });

  it('does not resolve unregistered formatters from Object.prototype', () => {
    const expression = compileTemplateExpression('toString(order.total)');

    expect(() =>
      evaluateTemplateExpression(expression, {
        root: { order: { total: 12.5 } },
      }),
    ).toThrowError(expect.objectContaining({ code: 'UNKNOWN_FORMATTER' }));
  });

  it('does not resolve formatters injected through Object.prototype pollution', () => {
    const formatterName = '__tegoPollutedFormatterForTest';
    const expression = compileTemplateExpression(`${formatterName}(order.total)`);
    const previous = Object.getOwnPropertyDescriptor(Object.prototype, formatterName);
    Object.defineProperty(Object.prototype, formatterName, {
      configurable: true,
      value: () => 'compromised',
    });

    try {
      expect(() =>
        evaluateTemplateExpression(expression, {
          root: { order: { total: 12.5 } },
        }),
      ).toThrowError(expect.objectContaining({ code: 'UNKNOWN_FORMATTER' }));
    } finally {
      if (previous === undefined)
        delete (Object.prototype as Record<string, unknown>)[formatterName];
      else Object.defineProperty(Object.prototype, formatterName, previous);
    }
  });

  it('exposes immutable repeat scope values', () => {
    const expression = compileTemplateExpression(
      'item.name + ":" + $index + ":" + $first + ":" + $last',
    );
    expect(
      evaluateTemplateExpression(expression, {
        root: {},
        item: { name: 'line' },
        index: 2,
        first: false,
        last: true,
      }),
    ).toBe('line:2:false:true');
  });
});
