import { describe, expect, it } from 'vitest';
import { createConditionalFormatEvaluator } from '../../../src/format/conditional';

describe('FMT-01 conditional formatting foundation', () => {
  it('applies ordered rules without mutating the base presentation', () => {
    const baseStyle = Object.freeze({ color: '#111111', bold: false });
    const evaluator = createConditionalFormatEvaluator({ maxRules: 10, maxCells: 100 });
    const result = evaluator.evaluate({
      address: { sheetId: 'sheet-1', row: 1, column: 0 },
      value: { type: 'number', value: 12 },
      text: '12',
      baseStyle,
      rules: [
        {
          id: 'high',
          priority: 1,
          stopIfTrue: true,
          ranges: [
            {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 9, column: 0 },
            },
          ],
          condition: { type: 'cell-is', operator: 'greaterThan', value: 10 },
          effect: { type: 'style', patch: { color: '#ff0000', bold: true } },
        },
        {
          id: 'later',
          priority: 2,
          stopIfTrue: false,
          ranges: [
            {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 9, column: 0 },
            },
          ],
          condition: { type: 'not-blank' },
          effect: { type: 'style', patch: { color: '#00ff00' } },
        },
      ],
    });

    expect(result).toMatchObject({
      matchedRuleIds: ['high'],
      stylePatch: { color: '#ff0000', bold: true },
      diagnostics: [],
    });
    expect(baseStyle).toEqual({ color: '#111111', bold: false });
  });

  it('rejects unsafe formulas and resource-limit overflows deterministically', () => {
    const evaluator = createConditionalFormatEvaluator({ maxRules: 1, maxCells: 1 });
    const input = {
      address: { sheetId: 'sheet-1', row: 0, column: 0 },
      value: { type: 'number' as const, value: 1 },
      text: '1',
      baseStyle: {},
    };
    expect(() =>
      evaluator.evaluate({
        ...input,
        rules: [
          {
            id: 'unsafe',
            priority: 1,
            stopIfTrue: false,
            ranges: [
              {
                sheetId: 'sheet-1',
                start: { row: 0, column: 0 },
                end: { row: 0, column: 0 },
              },
            ],
            condition: { type: 'formula', source: '=HYPERLINK("https://example.com")' },
            effect: { type: 'style', patch: { bold: true } },
          },
        ],
      }),
    ).toMatchObject({ code: 'INVALID_CONDITIONAL_EXPRESSION' });
    expect(() =>
      evaluator.evaluate({
        ...input,
        rules: [
          {
            id: 'one',
            priority: 1,
            stopIfTrue: false,
            ranges: [],
            condition: { type: 'blank' },
            effect: { type: 'style', patch: {} },
          },
          {
            id: 'two',
            priority: 2,
            stopIfTrue: false,
            ranges: [],
            condition: { type: 'blank' },
            effect: { type: 'style', patch: {} },
          },
        ],
      }),
    ).toMatchObject({ code: 'CONDITIONAL_RANGE_TOO_LARGE' });
  });
});
