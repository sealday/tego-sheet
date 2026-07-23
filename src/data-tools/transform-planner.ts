import type {
  DocumentCommandEnvelope,
  DocumentController,
  DocumentControllerSnapshot,
  DocumentTransactionEnvelope,
  DocumentTransactionResult,
} from '../document-controller';
import type { DocumentCellRange } from '../document';
import type { SheetId } from '../core';

export interface FindReplaceTransform {
  readonly type: 'find-replace';
  readonly range: DocumentCellRange;
  readonly find: string;
  readonly replacement: string;
  readonly match: 'literal';
}

export interface DataTransformPreview {
  readonly planId: string;
  readonly baseRevision: number;
  readonly sampleChanges: readonly {
    readonly row: number;
    readonly column: number;
    readonly before: string;
    readonly after: string;
  }[];
}

export type DataTransformCommitResult =
  | DocumentTransactionResult
  | {
      readonly status: 'rejected';
      readonly code: 'TRANSFORM_PLAN_MISSING' | 'TRANSFORM_PLAN_STALE';
    };

interface PendingPlan {
  readonly baseRevision: number;
  readonly commands: readonly DocumentCommandEnvelope[];
}

function inputText(
  input:
    | { readonly type: 'blank' }
    | { readonly type: 'string'; readonly value: string }
    | { readonly type: 'number'; readonly value: number }
    | { readonly type: 'boolean'; readonly value: boolean }
    | { readonly type: 'formula'; readonly source: string }
    | { readonly type: 'custom' },
): string {
  if (input.type === 'blank' || input.type === 'custom') return '';
  if (input.type === 'formula') return input.source;
  return String(input.value);
}

/** Creates an isolated revision-bound preview and commit planner. */
export function createDataTransformPlanner(limits: {
  readonly maxCells: number;
  readonly maxSamples: number;
}): {
  preview(
    snapshot: DocumentControllerSnapshot,
    transform: FindReplaceTransform,
  ): Promise<DataTransformPreview>;
  commit(controller: DocumentController, planId: string): DataTransformCommitResult;
} {
  const plans = new Map<string, PendingPlan>();
  let sequence = 0;
  return {
    async preview(snapshot, transform) {
      const rowCount = transform.range.end.row - transform.range.start.row + 1;
      const columnCount = transform.range.end.column - transform.range.start.column + 1;
      if (rowCount * columnCount > limits.maxCells) {
        throw new RangeError('Data transform exceeds the configured cell limit');
      }
      const sheet = snapshot.document.workbook.sheets.find(
        ({ id }) => id === transform.range.sheetId,
      );
      if (sheet === undefined) throw new RangeError('Data transform sheet does not exist');
      const changes: {
        row: number;
        column: number;
        before: string;
        after: string;
      }[] = [];
      const commands: DocumentCommandEnvelope[] = [];
      for (const entry of sheet.cells) {
        if (
          entry.row < transform.range.start.row ||
          entry.row > transform.range.end.row ||
          entry.column < transform.range.start.column ||
          entry.column > transform.range.end.column
        ) {
          continue;
        }
        const before = inputText(entry.cell.input);
        const after = before.replaceAll(transform.find, transform.replacement);
        if (before === after) continue;
        changes.push({ row: entry.row, column: entry.column, before, after });
        commands.push({
          schemaVersion: 1,
          id: `transform-cell-${commands.length + 1}`,
          command: {
            type: 'set-cell-text',
            address: {
              sheet: transform.range.sheetId as string as SheetId,
              row: entry.row,
              column: entry.column,
            },
            text: after,
          },
        });
      }
      const planId = `data-transform-${snapshot.revision}-${sequence}`;
      sequence += 1;
      plans.set(planId, { baseRevision: snapshot.revision, commands });
      return {
        planId,
        baseRevision: snapshot.revision,
        sampleChanges: Object.freeze(changes.slice(0, limits.maxSamples)),
      };
    },
    commit(controller, planId) {
      const plan = plans.get(planId);
      if (plan === undefined) return { status: 'rejected', code: 'TRANSFORM_PLAN_MISSING' };
      plans.delete(planId);
      if (controller.getSnapshot().revision !== plan.baseRevision) {
        return { status: 'rejected', code: 'TRANSFORM_PLAN_STALE' };
      }
      const transaction: DocumentTransactionEnvelope = {
        schemaVersion: 1,
        id: planId,
        baseRevision: plan.baseRevision,
        commands: plan.commands,
      };
      return controller.transact(transaction);
    },
  };
}
