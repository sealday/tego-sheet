import { describe, expect, it, vi } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import { createConditionalFormatEvaluator } from '../../../src/format/conditional';

describe('FMT-01 conditional formatting foundation', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;

  it('applies ordered rules without mutating the base presentation', () => {
    const baseStyle = Object.freeze({ color: '#111111', bold: false });
    const evaluator = createConditionalFormatEvaluator({ maxRules: 10, maxCells: 100 });
    const result = evaluator.evaluate({
      address: { sheetId, row: 1, column: 0 },
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
              sheetId,
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
              sheetId,
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
      address: { sheetId, row: 0, column: 0 },
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
                sheetId,
                start: { row: 0, column: 0 },
                end: { row: 0, column: 0 },
              },
            ],
            condition: { type: 'formula', source: '=HYPERLINK("https://example.com")' },
            effect: { type: 'style', patch: { bold: true } },
          },
        ],
      }),
    ).toThrowError(expect.objectContaining({ code: 'INVALID_CONDITIONAL_EXPRESSION' }));
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
    ).toThrowError(expect.objectContaining({ code: 'CONDITIONAL_RANGE_TOO_LARGE' }));
  });

  it('evaluates referenced formula expressions with synchronous allowlisted functions', () => {
    const evaluator = createConditionalFormatEvaluator({ maxRules: 10, maxCells: 100 });
    const result = evaluator.evaluate({
      address: { sheetId, row: 0, column: 2 },
      value: { type: 'number', value: 1 },
      text: '1',
      baseStyle: {},
      lookup: ({ row, column }) =>
        row === 0 && column === 0 ? { type: 'number', value: 12 } : { type: 'number', value: 3 },
      rules: [
        {
          id: 'formula',
          priority: 1,
          stopIfTrue: false,
          ranges: [{ sheetId, start: { row: 0, column: 2 }, end: { row: 0, column: 2 } }],
          condition: { type: 'formula', source: '=AND(A1>10,B1<5)' },
          effect: { type: 'style', patch: { bold: true } },
        },
      ],
    });

    expect(result).toMatchObject({
      matchedRuleIds: ['formula'],
      stylePatch: { bold: true },
      diagnostics: [],
    });
  });

  it('interpolates deterministic color-scale backgrounds from range values', () => {
    const evaluator = createConditionalFormatEvaluator({ maxRules: 10, maxCells: 100 });
    const values = [0, 5, 10];
    const result = evaluator.evaluate({
      address: { sheetId, row: 1, column: 0 },
      value: { type: 'number', value: 5 },
      text: '5',
      baseStyle: {},
      lookup: ({ row }) => ({ type: 'number', value: values[row] ?? 0 }),
      rules: [
        {
          id: 'scale',
          priority: 1,
          stopIfTrue: false,
          ranges: [{ sheetId, start: { row: 0, column: 0 }, end: { row: 2, column: 0 } }],
          condition: { type: 'not-blank' },
          effect: {
            type: 'color-scale',
            minimumColor: '#000000',
            maximumColor: '#ffffff',
          },
        },
      ],
    });

    expect(result).toMatchObject({
      matchedRuleIds: ['scale'],
      stylePatch: { backgroundColor: '#808080' },
      diagnostics: [],
    });
  });

  it('precomputes one color-scale range for repeated cell presentation', () => {
    const evaluator = createConditionalFormatEvaluator({ maxRules: 10, maxCells: 100 });
    const lookup = vi.fn(({ row }: { readonly row: number }) => ({
      type: 'number' as const,
      value: row,
    }));
    const rule = {
      id: 'scale-cache',
      priority: 1,
      stopIfTrue: false,
      ranges: [{ sheetId, start: { row: 0, column: 0 }, end: { row: 2, column: 0 } }],
      condition: { type: 'not-blank' as const },
      effect: {
        type: 'color-scale' as const,
        minimumColor: '#000000',
        maximumColor: '#ffffff',
      },
    };
    for (let row = 0; row < 3; row += 1) {
      evaluator.evaluate({
        address: { sheetId, row, column: 0 },
        value: { type: 'number', value: row },
        text: String(row),
        baseStyle: {},
        lookup,
        rules: [rule],
      });
    }

    expect(lookup).toHaveBeenCalledTimes(3);
  });

  it('enforces explicit formula source and evaluation budgets', () => {
    const evaluator = createConditionalFormatEvaluator({
      maxRules: 10,
      maxCells: 100,
      maxFormulaLength: 8,
      maxAstNodes: 4,
      maxEvaluationSteps: 4,
    });
    expect(() =>
      evaluator.evaluate({
        address: { sheetId, row: 0, column: 0 },
        value: { type: 'number', value: 1 },
        text: '1',
        baseStyle: {},
        rules: [
          {
            id: 'budget',
            priority: 1,
            stopIfTrue: false,
            ranges: [{ sheetId, start: { row: 0, column: 0 }, end: { row: 0, column: 0 } }],
            condition: { type: 'formula', source: '=1+1+1+1+1' },
            effect: { type: 'style', patch: { bold: true } },
          },
        ],
      }),
    ).toThrowError(expect.objectContaining({ code: 'INVALID_CONDITIONAL_EXPRESSION' }));
  });
});
