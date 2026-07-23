import { describe, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument, type DocumentSheetId } from '../../../src/document';
import {
  createValidationEngine,
  createValidationResolverRegistry,
  type ValidationRequest,
  type ValidationRule,
} from '../../../src/validation';
import { documentValidationRequest } from '../../../src/validation/document-rule';

const sheetId = 'sheet-1' as DocumentSheetId;

function request(rule: ValidationRule, value: ValidationRequest['value']): ValidationRequest {
  return {
    address: { sheetId, row: 0, column: 0 },
    rule,
    value,
  };
}

const base = {
  id: 'rule-1',
  behavior: 'reject' as const,
  allowBlank: false,
};

describe('advanced typed validation rules', () => {
  it.each([
    [
      'whole',
      {
        ...base,
        type: 'whole',
        predicate: { operator: 'between', minimum: 1, maximum: 3 },
      },
      { type: 'number', value: 2 },
      'accepted',
    ],
    [
      'whole rejects a fraction',
      {
        ...base,
        type: 'whole',
        predicate: { operator: 'between', minimum: 1, maximum: 3 },
      },
      { type: 'number', value: 2.5 },
      'rejected',
    ],
    [
      'decimal',
      {
        ...base,
        type: 'decimal',
        predicate: { operator: 'greaterThan', value: 1.25 },
      },
      { type: 'number', value: 1.5 },
      'accepted',
    ],
    [
      'date',
      {
        ...base,
        type: 'date',
        predicate: {
          operator: 'between',
          minimum: '2026-01-01',
          maximum: '2026-12-31',
        },
      },
      { type: 'string', value: '2026-07-24' },
      'accepted',
    ],
    [
      'time',
      {
        ...base,
        type: 'time',
        predicate: {
          operator: 'between',
          minimum: '09:00',
          maximum: '17:30:00',
        },
      },
      { type: 'string', value: '12:30' },
      'accepted',
    ],
    [
      'text length',
      {
        ...base,
        type: 'text-length',
        predicate: { operator: 'lessThanOrEqual', value: 5 },
      },
      { type: 'string', value: 'hello' },
      'accepted',
    ],
  ] as const)('evaluates %s rules', async (_name, rule, value, status) => {
    await expect(
      createValidationEngine().validate(
        request(rule as unknown as ValidationRule, value as ValidationRequest['value']),
      ),
    ).resolves.toMatchObject({ status });
  });

  it('evaluates custom formulas only through the injected bounded evaluator', async () => {
    const evaluateCustomFormula = vi.fn(() => true);
    const rule = {
      ...base,
      type: 'custom-formula',
      predicate: { formula: '=AND(A1<10,A1>0)' },
    } satisfies ValidationRule;
    const candidate = request(rule, { type: 'number', value: 5 });

    await expect(
      createValidationEngine({ evaluateCustomFormula }).validate(candidate),
    ).resolves.toEqual({ status: 'accepted', diagnostics: [] });
    expect(evaluateCustomFormula).toHaveBeenCalledWith(
      '=AND(A1<10,A1>0)',
      expect.objectContaining({
        request: candidate,
        signal: expect.any(AbortSignal),
      }),
    );
  });

  it.each([
    [
      {
        id: 'whole-rule',
        type: 'whole',
        predicate: { operator: 'equal', value: 7 },
        behavior: 'reject',
        allowBlank: false,
      },
      { type: 'number', value: 7 },
    ],
    [
      {
        id: 'date-rule',
        type: 'date',
        predicate: { operator: 'equal', value: '2026-07-24' },
        behavior: 'reject',
        allowBlank: false,
      },
      { type: 'string', value: '2026-07-24' },
    ],
  ] as const)('loads %s from the document validation registry', (rule, value) => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'validation-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              {
                row: 0,
                column: 0,
                cell: { input: { type: 'blank' }, validationId: rule.id },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [{ id: rule.id, value: rule }],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));

    expect(
      documentValidationRequest(
        parsed.document,
        { sheetId, row: 0, column: 0 },
        String('value' in value ? value.value : ''),
      ),
    ).toMatchObject({ rule, value });
  });

  it('coerces built-in checkbox and dropdown editor values to their typed scalars', () => {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'custom-validation-document',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: {
                    type: 'custom',
                    cellType: 'checkbox',
                    schemaVersion: 1,
                    value: { checked: false },
                  },
                  validationId: 'custom',
                },
              },
              {
                row: 0,
                column: 1,
                cell: {
                  input: {
                    type: 'custom',
                    cellType: 'dropdown',
                    schemaVersion: 1,
                    value: { value: 1, label: 'One' },
                  },
                  validationId: 'whole',
                },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [
          {
            id: 'custom',
            value: {
              id: 'custom',
              type: 'custom-formula',
              predicate: { formula: '=A1=TRUE' },
              behavior: 'reject',
              allowBlank: false,
            },
          },
          {
            id: 'whole',
            value: {
              id: 'whole',
              type: 'whole',
              predicate: { operator: 'equal', value: 7 },
              behavior: 'reject',
              allowBlank: false,
            },
          },
        ],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));

    expect(
      documentValidationRequest(parsed.document, { sheetId, row: 0, column: 0 }, 'true'),
    ).toMatchObject({ value: { type: 'boolean', value: true } });
    expect(
      documentValidationRequest(parsed.document, { sheetId, row: 0, column: 1 }, '7'),
    ).toMatchObject({ value: { type: 'number', value: 7 } });
  });
});

describe('resolver lifecycle limits', () => {
  it('hard-times out a resolver that ignores its abort signal', async () => {
    vi.useFakeTimers();
    try {
      const registry = createValidationResolverRegistry();
      registry.register('never', () => new Promise<readonly string[]>(() => {}));
      const engine = createValidationEngine({
        resolvers: registry,
        limits: { resolverTimeoutMs: 25 },
      });
      const result = engine.validate(
        request(
          {
            ...base,
            type: 'list',
            predicate: { source: { type: 'resolver', id: 'never' } },
          },
          { type: 'string', value: 'CN' },
        ),
      );

      await vi.advanceTimersByTimeAsync(25);
      await expect(result).resolves.toMatchObject({
        status: 'error',
        code: 'VALIDATION_SOURCE_TIMEOUT',
      });
    } finally {
      vi.useRealTimers();
    }
  });

  it('aborts in-flight validation when the caller aborts or the engine is disposed', async () => {
    const registry = createValidationResolverRegistry();
    const signals: AbortSignal[] = [];
    registry.register(
      'never',
      ({ signal }) =>
        new Promise<readonly string[]>((_resolve, reject) => {
          signals.push(signal);
          signal.addEventListener('abort', () => reject(signal.reason), {
            once: true,
          });
        }),
    );
    const engine = createValidationEngine({ resolvers: registry });
    const caller = new AbortController();
    const rule = {
      ...base,
      type: 'list',
      predicate: { source: { type: 'resolver', id: 'never' } },
    } satisfies ValidationRule;
    const first = engine.validate({
      ...request(rule, { type: 'string', value: 'CN' }),
      signal: caller.signal,
    });
    caller.abort();
    await expect(first).resolves.toMatchObject({
      status: 'error',
      code: 'VALIDATION_SOURCE_ABORTED',
    });

    const second = engine.validate(request(rule, { type: 'string', value: 'DE' }));
    engine.dispose?.();
    await expect(second).resolves.toMatchObject({
      status: 'error',
      code: 'VALIDATION_SOURCE_ABORTED',
    });
    expect(signals).toHaveLength(2);
    expect(signals.every((signal) => signal.aborted)).toBe(true);
  });
});
