import { describe, expect, it } from 'vitest';
import type { DocumentSheetId } from '../../../src/document';
import {
  parseSpreadsheetDocument,
  type JsonValue,
  type SpreadsheetDocument,
} from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';
import {
  createValidationEngine,
  createValidationResolverRegistry,
  executeValidatedCellEdit,
} from '../../../src/validation';

describe('VAL-01 validation foundation', () => {
  const sheetId = 'sheet-1' as DocumentSheetId;

  function documentWithRule(rule: JsonValue): SpreadsheetDocument {
    const parsed = parseSpreadsheetDocument({
      schemaVersion: 2,
      id: 'validation-edit',
      workbook: {
        sheets: [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            cells: [
              {
                row: 0,
                column: 0,
                cell: { input: { type: 'number', value: 1 }, validationId: 'amount' },
              },
            ],
            merges: [],
          },
        ],
        styles: [],
        validations: [{ id: 'amount', value: rule }],
        settings: { dateSystem: 'excel-1900' },
      },
      templates: [],
      resources: { items: [] },
      extensions: {},
    });
    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    return parsed.document;
  }

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
    const rule = {
      id: 'amount',
      type: 'number',
      predicate: { operator: 'between', minimum: 0, maximum: 10 },
      behavior: 'reject',
      allowBlank: false,
    };
    const controller = createDocumentController(documentWithRule(rule));

    await expect(
      executeValidatedCellEdit({
        controller,
        address: { sheetId, row: 0, column: 0 },
        text: '12',
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });
    expect(controller.getSnapshot().revision).toBe(0);

    const warningController = createDocumentController(
      documentWithRule({ ...rule, behavior: 'warn' }),
    );
    await expect(
      executeValidatedCellEdit({
        controller: warningController,
        address: { sheetId, row: 0, column: 0 },
        text: '12',
        confirmWarning: () => true,
      }),
    ).resolves.toMatchObject({ status: 'committed' });
    expect(warningController.getSnapshot().revision).toBe(1);
    expect(warningController.undo()).toMatchObject({ status: 'committed' });
  });

  it('rejects a stale async validation result instead of overwriting a newer revision', async () => {
    const controller = createDocumentController(
      documentWithRule({
        id: 'amount',
        type: 'list',
        predicate: { source: { type: 'resolver', id: 'delayed' } },
        behavior: 'reject',
        allowBlank: false,
      }),
    );
    let release!: (values: readonly string[]) => void;
    const values = new Promise<readonly string[]>((resolve) => {
      release = resolve;
    });
    const registry = createValidationResolverRegistry();
    registry.register('delayed', () => values);
    const pending = executeValidatedCellEdit({
      controller,
      address: { sheetId, row: 0, column: 0 },
      validation: { resolvers: registry },
      text: 'old',
    });
    controller.execute({
      schemaVersion: 1,
      id: 'newer-edit',
      command: {
        type: 'set-cell-text',
        address: { sheet: 'sheet-1' as never, row: 1, column: 0 },
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
