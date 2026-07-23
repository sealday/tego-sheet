import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import { createSpreadsheetDocument } from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';
import {
  createValidationEngine,
  createValidationResolverRegistry,
  executeValidatedCellEdit,
} from '../../../src/validation';

describe('VAL-01 validation foundation', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;

  it('distinguishes reject, warn, and accepted edits without mutation', async () => {
    const engine = createValidationEngine();
    const base = {
      address: { sheetId, row: 0, column: 0 },
      value: { type: 'number' as const, value: 12 },
    };
    await expect(
      engine.validate({
        ...base,
        rule: {
          id: 'amount',
          type: 'number',
          predicate: { operator: 'between', minimum: 0, maximum: 10 },
          behavior: 'reject',
          allowBlank: false,
        },
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });
    await expect(
      engine.validate({
        ...base,
        rule: {
          id: 'amount',
          type: 'number',
          predicate: { operator: 'between', minimum: 0, maximum: 10 },
          behavior: 'warn',
          allowBlank: false,
        },
      }),
    ).resolves.toMatchObject({ status: 'warning', code: 'VALIDATION_REJECTED' });
    await expect(
      engine.validate({
        ...base,
        value: { type: 'number', value: 5 },
        rule: {
          id: 'amount',
          type: 'number',
          predicate: { operator: 'between', minimum: 0, maximum: 10 },
          behavior: 'reject',
          allowBlank: false,
        },
      }),
    ).resolves.toEqual({ status: 'accepted', diagnostics: [] });
  });

  it('bounds and cancels host list resolvers without performing network access', async () => {
    const registry = createValidationResolverRegistry();
    registry.register('countries', async ({ signal }) => {
      await Promise.resolve();
      if (signal.aborted) throw signal.reason;
      return ['CN', 'NL', 'DE'];
    });
    const engine = createValidationEngine({
      resolvers: registry,
      limits: { maxListItems: 2, resolverTimeoutMs: 100 },
    });
    await expect(
      engine.validate({
        address: { sheetId, row: 0, column: 0 },
        value: { type: 'string', value: 'CN' },
        rule: {
          id: 'country',
          type: 'list',
          predicate: { source: { type: 'resolver', id: 'countries' } },
          behavior: 'reject',
          allowBlank: false,
        },
      }),
    ).resolves.toMatchObject({ status: 'error', code: 'VALIDATION_SOURCE_TOO_LARGE' });
  });

  it('commits accepted and confirmed warning edits through one document transaction', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' }),
    );
    const engine = createValidationEngine();
    const request = {
      address: { sheetId, row: 0, column: 0 },
      value: { type: 'number' as const, value: 12 },
      rule: {
        id: 'amount',
        type: 'number' as const,
        predicate: { operator: 'between' as const, minimum: 0, maximum: 10 },
        behavior: 'reject' as const,
        allowBlank: false,
      },
    };

    await expect(
      executeValidatedCellEdit({ controller, engine, request, text: '12' }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });
    expect(controller.getSnapshot().revision).toBe(0);

    await expect(
      executeValidatedCellEdit({
        controller,
        engine,
        request: { ...request, rule: { ...request.rule, behavior: 'warn' } },
        text: '12',
        confirmWarning: () => true,
      }),
    ).resolves.toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().revision).toBe(1);
    expect(controller.undo()).toMatchObject({ status: 'committed' });
  });
});
