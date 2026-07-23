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

  it('rejects a stale async validation result instead of overwriting a newer revision', async () => {
    const controller = createDocumentController(
      createSpreadsheetDocument({ id: 'document-race', sheetId: 'sheet-1' }),
    );
    let release!: (values: readonly string[]) => void;
    const values = new Promise<readonly string[]>((resolve) => {
      release = resolve;
    });
    const registry = createValidationResolverRegistry();
    registry.register('delayed', () => values);
    const pending = executeValidatedCellEdit({
      controller,
      engine: createValidationEngine({ resolvers: registry }),
      request: {
        address: { sheetId, row: 0, column: 0 },
        value: { type: 'string', value: 'old' },
        rule: {
          id: 'delayed-rule',
          type: 'list',
          predicate: { source: { type: 'resolver', id: 'delayed' } },
          behavior: 'reject',
          allowBlank: false,
        },
      },
      text: 'old',
    });
    controller.execute({
      schemaVersion: 1,
      id: 'newer-edit',
      command: {
        type: 'set-cell-text',
        address: { sheet: 'sheet-1' as never, row: 0, column: 0 },
        text: 'newer',
      },
    });
    release(['old']);

    await expect(pending).resolves.toMatchObject({
      status: 'rejected',
      code: 'REVISION_CONFLICT',
    });
    expect(controller.getSnapshot().revision).toBe(1);
  });
});
