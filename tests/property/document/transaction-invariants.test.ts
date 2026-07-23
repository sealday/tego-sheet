import { describe, expect, it, vi } from 'vitest';
import { sheetId } from '../../../src/core';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
} from '../../../src/document';
import {
  SpreadsheetDocumentController,
  type SerializableCommandEnvelope,
  type SerializableTransactionEnvelope,
} from '../../../src/core/controller/spreadsheet-document-controller';

function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state * 1_664_525 + 1_013_904_223) >>> 0;
    return state;
  };
}

function createMultiSheetDocument(seed: number) {
  const input = JSON.parse(
    serializeSpreadsheetDocument(
      createSpreadsheetDocument({ id: `document-${seed}`, sheetId: 'sheet-1' }),
    ),
  ) as {
    workbook: {
      sheets: Array<{ id: string; name: string; cells: unknown[]; merges: unknown[] }>;
    };
  };
  input.workbook.sheets.push({ id: 'sheet-2', name: 'Second', cells: [], merges: [] });
  const parsed = parseSpreadsheetDocument(input as never);
  if (!parsed.ok) throw new TypeError('Expected generated document to parse');
  return parsed.document;
}

function setText(
  id: string,
  sheet: 'sheet-1' | 'sheet-2',
  row: number,
  column: number,
  text: string,
): SerializableCommandEnvelope {
  return {
    schemaVersion: 1,
    id,
    command: {
      type: 'set-cell-text',
      address: { sheet: sheetId(sheet), row, column },
      text,
    },
  };
}

describe('document transaction generated invariants', () => {
  it('preserves atomicity, inverse history, branching, serialization, and rejection invariants', () => {
    for (let seed = 1; seed <= 25; seed += 1) {
      const controller = new SpreadsheetDocumentController(createMultiSheetDocument(seed));
      const next = random(seed);
      const dimensions = {
        'sheet-1': { rows: 100, columns: 26 },
        'sheet-2': { rows: 100, columns: 26 },
      };
      const listener = vi.fn();
      controller.subscribe(listener);

      const initialBytes = serializeSpreadsheetDocument(controller.getDocument());
      const initialRevision = controller.getSnapshot().revision;
      const initialHistory = controller.historySize;
      const rejected: SerializableTransactionEnvelope = {
        schemaVersion: 1,
        id: `rejected-${seed}`,
        baseRevision: 99,
        commands: [setText(`rejected-command-${seed}`, 'sheet-1', 0, 0, 'rejected')],
      };
      expect(controller.transact(rejected)).toMatchObject({
        status: 'rejected',
        code: 'REVISION_CONFLICT',
      });
      expect(
        controller.transact(
          { ...rejected, baseRevision: initialRevision },
          { permissionGate: () => false },
        ),
      ).toMatchObject({ status: 'rejected', code: 'COMMAND_NOT_ALLOWED' });
      expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(initialBytes);
      expect(controller.getSnapshot().revision).toBe(initialRevision);
      expect(controller.historySize).toEqual(initialHistory);

      for (let transactionIndex = 0; transactionIndex < 3; transactionIndex += 1) {
        const commands: SerializableCommandEnvelope[] = [
          setText(
            `command-${seed}-${transactionIndex}-0`,
            transactionIndex % 2 === 0 ? 'sheet-1' : 'sheet-2',
            next() % 12,
            next() % 8,
            `value-${seed}-${transactionIndex}`,
          ),
        ];
        for (let commandIndex = 1; commandIndex < 9; commandIndex += 1) {
          const sheet = next() % 2 === 0 ? 'sheet-1' : 'sheet-2';
          const axis = next() % 2 === 0 ? 'row' : 'column';
          const insert = next() % 3 === 0;
          const dimension = dimensions[sheet];
          const size = axis === 'row' ? dimension.rows : dimension.columns;
          const index = next() % Math.min(Math.max(1, size - 1), 12);
          commands.push({
            schemaVersion: 1,
            id: `command-${seed}-${transactionIndex}-${commandIndex}`,
            command: {
              type: `${insert ? 'insert' : 'delete'}-${axis}`,
              sheet: sheetId(sheet),
              index,
              count: 1,
            },
          });
          if (axis === 'row') dimension.rows += insert ? 1 : -1;
          else dimension.columns += insert ? 1 : -1;
        }
        commands.push(
          setText(
            `command-${seed}-${transactionIndex}-9`,
            'sheet-2',
            20,
            10,
            `terminal-${seed}-${transactionIndex}`,
          ),
        );
        const envelope = JSON.parse(
          JSON.stringify({
            schemaVersion: 1,
            id: `transaction-${seed}-${transactionIndex}`,
            baseRevision: controller.getSnapshot().revision,
            commands,
            metadata: { seed, transactionIndex },
          }),
        ) as SerializableTransactionEnvelope;
        const before = serializeSpreadsheetDocument(controller.getDocument());
        const preview = controller.dryRun(envelope);
        expect(preview.status, `preview ${seed}:${transactionIndex}`).toBe('ready');
        expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(before);

        const outcome = controller.transact(envelope);
        expect(outcome.status, `commit ${seed}:${transactionIndex}`).toBe('committed');
        expect(controller.validate().valid, `invariant ${seed}:${transactionIndex}`).toBe(true);
        const after = serializeSpreadsheetDocument(controller.getDocument());
        if (preview.status === 'ready') {
          expect(serializeSpreadsheetDocument(preview.document)).toBe(after);
        }
        expect(controller.undo().status, `undo ${seed}:${transactionIndex}`).toBe('committed');
        expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(before);
        expect(controller.redo().status, `redo ${seed}:${transactionIndex}`).toBe('committed');
        expect(serializeSpreadsheetDocument(controller.getDocument())).toBe(after);
      }

      expect(controller.undo().status, `branch undo ${seed}`).toBe('committed');
      expect(
        controller.execute(setText(`branch-${seed}`, 'sheet-2', 20, 10, `branch-${seed}`)).status,
        `branch commit ${seed}`,
      ).toBe('committed');
      expect(controller.redo().status, `branch redo ${seed}`).toBe('noop');
      expect(controller.historySize.redo, `branch history ${seed}`).toBe(0);
      expect(controller.validate().valid, `branch invariant ${seed}`).toBe(true);
      expect(listener.mock.calls.length, `event count ${seed}`).toBeGreaterThanOrEqual(10);
    }
  });
});
