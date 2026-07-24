import { describe, expect, it } from 'vitest';
import type { SheetId } from '../../../src/core';
import {
  issueValidationCapability,
  SpreadsheetDocumentController,
} from '../../../src/core/controller/spreadsheet-document-controller';
import { parseSpreadsheetDocument, type SpreadsheetDocument } from '../../../src/document';
import { createDocumentController } from '../../../src/document-controller';
import {
  createValidationResolverRegistry,
  executeValidatedCellEdit,
  executeValidatedTransaction,
} from '../../../src/validation';

function document(rule: unknown): SpreadsheetDocument {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'validation-boundary',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'number', value: 1 }, validationId: 'rule-1' },
            },
          ],
          merges: [],
        },
      ],
      styles: [],
      validations: [{ id: 'rule-1', value: rule }],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return parsed.document;
}

const wholeRule = {
  id: 'rule-1',
  type: 'number',
  predicate: { operator: 'between', minimum: 0, maximum: 10 },
  behavior: 'reject',
  allowBlank: false,
} as const;

const address = { sheet: 'sheet-1' as SheetId, row: 0, column: 0 };

describe('validation mutation boundary', () => {
  it('keeps validation capabilities off the public controller surface', () => {
    const controller = createDocumentController(document(wholeRule));
    expect('issueValidationCapability' in controller).toBe(false);
    expect('transactValidated' in controller).toBe(false);
  });

  it('sync-validates static rules and rejects async or malformed rules with stable codes', () => {
    const staticController = new SpreadsheetDocumentController(document(wholeRule));
    expect(
      staticController.dispatch({ type: 'set-cell-text', address, text: '5' }, 'ref'),
    ).toMatchObject({ status: 'committed' });
    expect(
      staticController.execute({
        schemaVersion: 1,
        id: 'invalid-static',
        command: { type: 'set-cell-text', address, text: '50' },
      }),
    ).toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });

    const asyncController = new SpreadsheetDocumentController(
      document({
        ...wholeRule,
        type: 'list',
        predicate: { source: { type: 'resolver', id: 'remote' } },
      }),
    );
    expect(
      asyncController.execute({
        schemaVersion: 1,
        id: 'async-direct',
        command: { type: 'set-cell-text', address, text: 'allowed' },
      }),
    ).toMatchObject({ status: 'rejected', code: 'ASYNC_REQUIRED' });

    const malformed = new SpreadsheetDocumentController(
      document({ ...wholeRule, predicate: { operator: 'between', minimum: 'bad', maximum: 10 } }),
    );
    expect(
      malformed.execute({
        schemaVersion: 1,
        id: 'malformed-direct',
        command: { type: 'set-cell-text', address, text: '5' },
      }),
    ).toMatchObject({ status: 'rejected', code: 'VALIDATION_RULE_INVALID' });
  });

  it('consumes an unforgeable validation capability once', () => {
    const controller = new SpreadsheetDocumentController(
      document({
        ...wholeRule,
        type: 'list',
        predicate: { source: { type: 'resolver', id: 'remote' } },
      }),
    );
    const transaction = {
      schemaVersion: 1 as const,
      id: 'validated',
      baseRevision: 0,
      commands: [
        {
          schemaVersion: 1 as const,
          id: 'cell',
          command: { type: 'set-cell-text' as const, address, text: 'allowed' },
        },
      ],
    };
    expect(controller.transactValidated(transaction, {}, {})).toMatchObject({
      status: 'rejected',
      code: 'VALIDATION_CAPABILITY_INVALID',
    });
    const capability = issueValidationCapability(controller, transaction);
    expect(
      controller.transactValidated(
        {
          ...transaction,
          commands: [
            {
              ...transaction.commands[0]!,
              command: { type: 'set-cell-text', address, text: 'forged' },
            },
          ],
        },
        {},
        capability,
      ),
    ).toMatchObject({ status: 'rejected', code: 'VALIDATION_CAPABILITY_INVALID' });
    expect(controller.transactValidated(transaction, {}, capability)).toMatchObject({
      status: 'committed',
    });
    expect(
      controller.transactValidated({ ...transaction, baseRevision: 1 }, {}, capability),
    ).toMatchObject({ status: 'rejected', code: 'VALIDATION_CAPABILITY_INVALID' });
  });

  it('validates a paste batch before one transaction and preserves one undo entry', async () => {
    let release!: (values: readonly string[]) => void;
    const values = new Promise<readonly string[]>((resolve) => {
      release = resolve;
    });
    const resolvers = createValidationResolverRegistry();
    resolvers.register('remote', () => values);
    const controller = createDocumentController(
      document({
        ...wholeRule,
        type: 'list',
        predicate: { source: { type: 'resolver', id: 'remote' } },
      }),
    );
    const otherController = createDocumentController(
      document({
        ...wholeRule,
        type: 'list',
        predicate: { source: { type: 'resolver', id: 'remote' } },
      }),
    );
    const transaction = {
      schemaVersion: 1 as const,
      id: 'paste',
      baseRevision: 0,
      commands: [
        {
          schemaVersion: 1 as const,
          id: 'paste-values',
          command: {
            type: 'paste-external' as const,
            target: {
              sheet: 'sheet-1' as SheetId,
              active: { row: 0, column: 0 },
              range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
            },
            values: [['allowed']],
          },
        },
      ],
    };
    const validationInput = {
      controller,
      validation: { resolvers },
      transaction,
    };
    const pending = executeValidatedTransaction(validationInput);
    expect(controller.getSnapshot().revision).toBe(0);
    validationInput.controller = otherController;
    transaction.commands[0]!.command.values[0]![0] = 'forbidden';
    release(['allowed']);
    await expect(pending).resolves.toMatchObject({ status: 'committed', revision: 1 });
    expect(
      controller
        .getSnapshot()
        .document.workbook.sheets[0]!.cells.find(({ row, column }) => row === 0 && column === 0)
        ?.cell.input,
    ).toEqual({ type: 'string', value: 'allowed' });
    expect(otherController.getSnapshot().revision).toBe(0);
    expect(controller.undo()).toMatchObject({ status: 'committed', revision: 2 });
  });

  it('rejects hostile transaction accessors with a stable schema code', async () => {
    const controller = createDocumentController(document(wholeRule));
    const transaction = {
      schemaVersion: 1,
      id: 'hostile',
      commands: [],
    } as Record<string, unknown>;
    let getterCalls = 0;
    Object.defineProperty(transaction, 'baseRevision', {
      enumerable: true,
      get() {
        getterCalls += 1;
        throw new Error('must not execute');
      },
    });
    await expect(
      executeValidatedTransaction({
        controller,
        transaction: transaction as never,
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'COMMAND_SCHEMA_INVALID' });
    expect(getterCalls).toBe(0);
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('does not execute hostile validation request getters', async () => {
    const controller = createDocumentController(document(wholeRule));
    let getterCalls = 0;
    const input = {
      transaction: {
        schemaVersion: 1,
        id: 'hostile-request',
        baseRevision: 0,
        commands: [],
      },
    } as Record<string, unknown>;
    Object.defineProperty(input, 'controller', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return controller;
      },
    });
    await expect(executeValidatedTransaction(input as never)).resolves.toMatchObject({
      status: 'rejected',
      code: 'COMMAND_SCHEMA_INVALID',
    });
    expect(getterCalls).toBe(0);
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('rejects clear batches atomically when blank violates any rule', async () => {
    const controller = createDocumentController(document(wholeRule));
    await expect(
      executeValidatedTransaction({
        controller,
        transaction: {
          schemaVersion: 1,
          id: 'clear',
          baseRevision: 0,
          commands: [
            {
              schemaVersion: 1,
              id: 'clear-values',
              command: {
                type: 'clear-contents',
                selection: {
                  sheet: 'sheet-1' as SheetId,
                  active: { row: 0, column: 0 },
                  range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
                },
              },
            },
          ],
        },
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('rejects oversized clear batches before invoking the validation engine', async () => {
    const controller = createDocumentController(document(wholeRule));
    await expect(
      executeValidatedTransaction({
        controller,
        transaction: {
          schemaVersion: 1,
          id: 'oversized-clear',
          baseRevision: 0,
          commands: [
            {
              schemaVersion: 1,
              id: 'clear-values',
              command: {
                type: 'clear-contents',
                selection: {
                  sheet: 'sheet-1' as SheetId,
                  active: { row: 0, column: 0 },
                  range: {
                    start: { row: 0, column: 0 },
                    end: { row: Number.MAX_SAFE_INTEGER, column: 1 },
                  },
                },
              },
            },
          ],
        },
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'TRANSACTION_LIMIT_EXCEEDED' });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('ignores forged caller engines and rules in favor of the document-owned rule', async () => {
    const controller = createDocumentController(document(wholeRule));
    const forged = {
      controller,
      address: {
        sheetId: 'sheet-1' as import('../../../src/document').DocumentSheetId,
        row: 0,
        column: 0,
      },
      text: '50',
      engine: { validate: async () => ({ status: 'accepted', diagnostics: [] }) },
      request: {
        value: { type: 'number', value: 5 },
        rule: { ...wholeRule, predicate: { operator: 'between', minimum: 0, maximum: 100 } },
      },
    } as unknown as Parameters<typeof executeValidatedCellEdit>[0];
    await expect(executeValidatedCellEdit(forged)).resolves.toMatchObject({
      status: 'rejected',
      code: 'VALIDATION_REJECTED',
    });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it.each(['paste-internal', 'autofill'] as const)(
    'validates values projected by %s into an existing target rule',
    async (type) => {
      const parsed = parseSpreadsheetDocument({
        ...document(wholeRule),
        workbook: {
          ...document(wholeRule).workbook,
          sheets: [
            {
              ...document(wholeRule).workbook.sheets[0]!,
              cells: [
                {
                  row: 0,
                  column: 0,
                  cell: { input: { type: 'number', value: 1 }, validationId: 'rule-1' },
                },
                { row: 0, column: 1, cell: { input: { type: 'number', value: 50 } } },
              ],
            },
          ],
        },
      });
      if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
      const controller = createDocumentController(parsed.document);
      const selection = (column: number) => ({
        sheet: 'sheet-1' as SheetId,
        active: { row: 0, column },
        range: { start: { row: 0, column }, end: { row: 0, column } },
      });
      const command =
        type === 'paste-internal'
          ? {
              type,
              source: selection(1),
              target: selection(0),
              mode: 'value' as const,
              cut: false,
            }
          : {
              type,
              source: selection(1),
              target: selection(0),
              mode: 'value' as const,
            };
      await expect(
        executeValidatedTransaction({
          controller,
          transaction: {
            schemaVersion: 1,
            id: type,
            baseRevision: 0,
            commands: [
              {
                schemaVersion: 1,
                id: type,
                command,
              },
            ],
          },
        }),
      ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_REJECTED' });
      expect(controller.getSnapshot().revision).toBe(0);
    },
  );

  it('rejects invalid batch limits before preview or validation work', async () => {
    const controller = createDocumentController(document(wholeRule));
    await expect(
      executeValidatedTransaction({
        controller,
        limits: { maxConcurrency: Number.NaN, totalTimeoutMs: 10 },
        transaction: {
          schemaVersion: 1,
          id: 'invalid-limits',
          baseRevision: 0,
          commands: [],
        },
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_LIMIT_INVALID' });
  });

  it('applies the total deadline to warning confirmation', async () => {
    const controller = createDocumentController(document({ ...wholeRule, behavior: 'warn' }));
    await expect(
      executeValidatedTransaction({
        controller,
        limits: { totalTimeoutMs: 10 },
        confirmWarning: () => new Promise<boolean>(() => undefined),
        transaction: {
          schemaVersion: 1,
          id: 'warning-timeout',
          baseRevision: 0,
          commands: [
            {
              schemaVersion: 1,
              id: 'warning-edit',
              command: { type: 'set-cell-text', address, text: '50' },
            },
          ],
        },
      }),
    ).resolves.toMatchObject({ status: 'rejected', code: 'VALIDATION_SOURCE_TIMEOUT' });
    expect(controller.getSnapshot().revision).toBe(0);
  });

  it('keeps advanced typed rules out of legacy runtime registries after a committed clear', async () => {
    const controller = createDocumentController(document({ ...wholeRule, allowBlank: true }));
    await expect(
      executeValidatedTransaction({
        controller,
        transaction: {
          schemaVersion: 1,
          id: 'clear-advanced-rule',
          baseRevision: 0,
          commands: [
            {
              schemaVersion: 1,
              id: 'clear-cell',
              command: {
                type: 'clear-contents',
                selection: {
                  sheet: 'sheet-1' as SheetId,
                  active: { row: 0, column: 0 },
                  range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
                },
              },
            },
          ],
        },
      }),
    ).resolves.toMatchObject({ status: 'committed' });
    expect(controller.getSnapshot().document.workbook.validations).toEqual([
      expect.objectContaining({ id: 'rule-1' }),
    ]);
    expect(
      controller
        .getSnapshot()
        .document.workbook.validations.some(({ id }) => id.startsWith('runtime-validation-')),
    ).toBe(false);
  });
});
