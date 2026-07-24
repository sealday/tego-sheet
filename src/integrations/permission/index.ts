import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type {
  TransactionPermissionContext,
  TransactionPermissionGate,
} from '../../core/controller/spreadsheet-document-controller';

/** Stable permission actions shared by every product entry point. */
export type PermissionAction =
  | 'document:view'
  | 'document:edit'
  | 'sheet:view'
  | 'sheet:edit'
  | 'range:edit'
  | 'object:edit'
  | 'comment:view'
  | 'comment:create'
  | 'comment:resolve'
  | 'template:bind'
  | 'print'
  | 'download'
  | 'history:view'
  | 'history:restore'
  | 'ai:propose'
  | 'ai:apply';

export interface PermissionRange {
  readonly sheetId: string;
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
}

/** Resource target checked by one permission action. */
export type PermissionTarget =
  | { readonly type: 'document'; readonly documentId: string }
  | { readonly type: 'sheet'; readonly sheetId: string }
  | { readonly type: 'range'; readonly range: PermissionRange }
  | { readonly type: 'object'; readonly sheetId: string; readonly objectId: string }
  | { readonly type: 'comment'; readonly threadId: string };

export interface PermissionGrant {
  readonly action: PermissionAction;
  readonly target: PermissionTarget;
}

export interface CreatePermissionSnapshotOptions {
  readonly revision: string;
  readonly actorId: string;
  readonly grants: readonly PermissionGrant[];
}

/** Immutable revisioned permission snapshot. Client checks never replace server authorization. */
export interface PermissionSnapshot {
  readonly revision: string;
  readonly actorId: string;
  can(action: PermissionAction, target: PermissionTarget): boolean;
}

export interface PermissionRequest {
  readonly action: PermissionAction;
  readonly target: PermissionTarget;
}

export interface PermissionDecision {
  readonly allowed: boolean;
  readonly snapshotRevision: string;
  readonly deniedTargets: readonly PermissionTarget[];
}

export interface EvaluatePermissionOptions {
  readonly expectedRevision?: string;
  readonly maximumTargets?: number;
}

export interface PermissionStore {
  getSnapshot(): PermissionSnapshot | undefined;
  replace(snapshot: PermissionSnapshot): void;
  clear(): void;
  can(action: PermissionAction, target: PermissionTarget): boolean;
  subscribe(listener: (snapshot: PermissionSnapshot | undefined) => void): () => void;
}

export interface CreateWorkbookTransactionPermissionGateOptions {
  readonly documentId: string;
  readonly getSnapshot: () => PermissionSnapshot | undefined;
}

const actions = new Set<PermissionAction>([
  'document:view',
  'document:edit',
  'sheet:view',
  'sheet:edit',
  'range:edit',
  'object:edit',
  'comment:view',
  'comment:create',
  'comment:resolve',
  'template:bind',
  'print',
  'download',
  'history:view',
  'history:restore',
  'ai:propose',
  'ai:apply',
]);
const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function coordinate(value: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new TypeError(`${label} must be a non-negative safe integer`);
  }
  return value;
}

function snapshotRange(range: PermissionRange): PermissionRange {
  const startRow = coordinate(Math.min(range.start.row, range.end.row), 'Permission range row');
  const endRow = coordinate(Math.max(range.start.row, range.end.row), 'Permission range row');
  const startColumn = coordinate(
    Math.min(range.start.column, range.end.column),
    'Permission range column',
  );
  const endColumn = coordinate(
    Math.max(range.start.column, range.end.column),
    'Permission range column',
  );
  return Object.freeze({
    sheetId: identifier(range.sheetId, 'Permission range sheetId'),
    start: Object.freeze({ row: startRow, column: startColumn }),
    end: Object.freeze({ row: endRow, column: endColumn }),
  });
}

function snapshotTarget(target: PermissionTarget): PermissionTarget {
  if (target.type === 'document') {
    return Object.freeze({
      type: 'document',
      documentId: identifier(target.documentId, 'Permission documentId'),
    });
  }
  if (target.type === 'sheet') {
    return Object.freeze({
      type: 'sheet',
      sheetId: identifier(target.sheetId, 'Permission sheetId'),
    });
  }
  if (target.type === 'range') {
    return Object.freeze({ type: 'range', range: snapshotRange(target.range) });
  }
  if (target.type === 'object') {
    return Object.freeze({
      type: 'object',
      sheetId: identifier(target.sheetId, 'Permission object sheetId'),
      objectId: identifier(target.objectId, 'Permission objectId'),
    });
  }
  return Object.freeze({
    type: 'comment',
    threadId: identifier(target.threadId, 'Permission comment threadId'),
  });
}

function targetMatches(granted: PermissionTarget, requested: PermissionTarget): boolean {
  if (granted.type !== requested.type) return false;
  if (granted.type === 'document' && requested.type === 'document') {
    return granted.documentId === requested.documentId;
  }
  if (granted.type === 'sheet' && requested.type === 'sheet') {
    return granted.sheetId === requested.sheetId;
  }
  if (granted.type === 'object' && requested.type === 'object') {
    return granted.sheetId === requested.sheetId && granted.objectId === requested.objectId;
  }
  if (granted.type === 'comment' && requested.type === 'comment') {
    return granted.threadId === requested.threadId;
  }
  if (granted.type === 'range' && requested.type === 'range') {
    const allowed = granted.range;
    const target = snapshotRange(requested.range);
    return (
      allowed.sheetId === target.sheetId &&
      allowed.start.row <= target.start.row &&
      allowed.start.column <= target.start.column &&
      allowed.end.row >= target.end.row &&
      allowed.end.column >= target.end.column
    );
  }
  return false;
}

function rangeRequest(
  sheetId: string,
  start: { readonly row: number; readonly column: number },
  end = start,
): PermissionRequest {
  return {
    action: 'range:edit',
    target: {
      type: 'range',
      range: {
        sheetId,
        start: { row: start.row, column: start.column },
        end: { row: end.row, column: end.column },
      },
    },
  };
}

function sheetViewRequest(sheetId: string): PermissionRequest {
  return { action: 'sheet:view', target: { type: 'sheet', sheetId } };
}

/** Derives the document guard and precise mutation target for one workbook command. */
export function deriveWorkbookCommandPermissionRequests(
  command: WorkbookCommand,
  documentId: string,
): readonly PermissionRequest[] {
  const documentRequest: PermissionRequest = {
    action: 'document:edit',
    target: { type: 'document', documentId: identifier(documentId, 'Permission documentId') },
  };
  let targetRequest: PermissionRequest;
  switch (command.type) {
    case 'set-cell-text':
    case 'set-cell-input':
      targetRequest = rangeRequest(command.address.sheet, command.address);
      break;
    case 'clear-contents':
    case 'set-cell-metadata':
    case 'set-style':
    case 'set-border':
    case 'clear-format':
    case 'merge':
    case 'unmerge':
    case 'set-validation':
    case 'remove-validation':
    case 'set-filter':
      targetRequest = rangeRequest(
        command.selection.sheet,
        command.selection.range.start,
        command.selection.range.end,
      );
      break;
    case 'paste-external':
      targetRequest = rangeRequest(
        command.target.sheet,
        command.target.range.start,
        command.target.range.end,
      );
      break;
    case 'paint-format':
    case 'autofill':
      return Object.freeze([
        documentRequest,
        rangeRequest(command.target.sheet, command.target.range.start, command.target.range.end),
        sheetViewRequest(command.source.sheet),
      ]);
    case 'paste-internal': {
      const target = rangeRequest(
        command.target.sheet,
        command.target.range.start,
        command.target.range.end,
      );
      const sourceView = sheetViewRequest(command.source.sheet);
      if (!command.cut) return Object.freeze([documentRequest, target, sourceView]);
      return Object.freeze([
        documentRequest,
        target,
        sourceView,
        rangeRequest(command.source.sheet, command.source.range.start, command.source.range.end),
      ]);
    }
    case 'set-sheet-object':
      targetRequest = {
        action: 'object:edit',
        target: { type: 'object', sheetId: command.sheet, objectId: command.object.id },
      };
      break;
    case 'remove-sheet-object':
      targetRequest = {
        action: 'object:edit',
        target: { type: 'object', sheetId: command.sheet, objectId: command.objectId },
      };
      break;
    case 'set-chart': {
      const sourceSheets = new Set<string>();
      if (command.chart.categories !== undefined) {
        sourceSheets.add(command.chart.categories.sheetId);
      }
      for (const series of command.chart.series) {
        sourceSheets.add(series.values.sheetId);
      }
      return Object.freeze([
        documentRequest,
        {
          action: 'object:edit',
          target: { type: 'object', sheetId: command.sheet, objectId: command.chart.id },
        },
        ...Array.from(sourceSheets, sheetViewRequest),
      ]);
    }
    case 'remove-chart':
      targetRequest = {
        action: 'object:edit',
        target: { type: 'object', sheetId: command.sheet, objectId: command.chartId },
      };
      break;
    case 'set-sparkline':
      return Object.freeze([
        documentRequest,
        rangeRequest(command.sparkline.target.sheetId, command.sparkline.target),
        sheetViewRequest(command.sparkline.source.sheetId),
      ]);
    case 'remove-sparkline':
      targetRequest = {
        action: 'sheet:edit',
        target: { type: 'sheet', sheetId: command.sheet },
      };
      break;
    case 'add-sheet':
    case 'undo':
    case 'redo':
      return Object.freeze([documentRequest]);
    case 'insert-row':
    case 'delete-row':
    case 'insert-column':
    case 'delete-column':
    case 'set-row-height':
    case 'set-row-hidden':
    case 'set-column-width':
    case 'set-column-hidden':
    case 'set-freeze':
    case 'delete-sheet':
    case 'rename-sheet':
    case 'group':
    case 'ungroup':
    case 'toggle-group':
    case 'clear-filter':
    case 'sort':
    case 'set-conditional-format':
    case 'remove-conditional-format':
    case 'set-filter-view':
    case 'remove-filter-view':
    case 'set-table':
    case 'remove-table':
      targetRequest = {
        action: 'sheet:edit',
        target: { type: 'sheet', sheetId: command.sheet },
      };
      break;
  }
  return Object.freeze([documentRequest, targetRequest]);
}

/** Creates one atomic default-deny gate for public controller transactions. */
export function createWorkbookTransactionPermissionGate(
  options: CreateWorkbookTransactionPermissionGateOptions,
): TransactionPermissionGate {
  const documentId = identifier(options.documentId, 'Permission documentId');
  return (context: TransactionPermissionContext): boolean => {
    const requests = context.transaction.commands.flatMap(({ command }) =>
      deriveWorkbookCommandPermissionRequests(command, documentId),
    );
    return evaluatePermission(options.getSnapshot(), requests).allowed;
  };
}

/** Creates an immutable, normalized, default-deny permission snapshot. */
export function createPermissionSnapshot(
  options: CreatePermissionSnapshotOptions,
): PermissionSnapshot {
  const revision = identifier(options.revision, 'Permission revision');
  const actorId = identifier(options.actorId, 'Permission actorId');
  if (options.grants.length > 100_000) throw new RangeError('Permission grant limit exceeded');
  const grants = Object.freeze(
    options.grants.map((grant) => {
      if (!actions.has(grant.action))
        throw new TypeError(`Unknown permission action ${grant.action}`);
      return Object.freeze({ action: grant.action, target: snapshotTarget(grant.target) });
    }),
  );
  return Object.freeze({
    revision,
    actorId,
    can(action: PermissionAction, target: PermissionTarget): boolean {
      if (!actions.has(action)) return false;
      let normalized: PermissionTarget;
      try {
        normalized = snapshotTarget(target);
      } catch {
        return false;
      }
      return grants.some(
        (grant) => grant.action === action && targetMatches(grant.target, normalized),
      );
    },
  });
}

/** Evaluates every target against one revision and never partially authorizes an operation. */
export function evaluatePermission(
  snapshot: PermissionSnapshot | undefined,
  requests: readonly PermissionRequest[],
  options: EvaluatePermissionOptions = {},
): PermissionDecision {
  const maximumTargets = options.maximumTargets ?? 10_000;
  if (!Number.isSafeInteger(maximumTargets) || maximumTargets < 1 || maximumTargets > 10_000) {
    throw new RangeError('Permission maximumTargets must be from 1 through 10000');
  }
  if (requests.length > maximumTargets) {
    throw new RangeError(`Permission operation cannot exceed ${maximumTargets} discrete targets`);
  }
  if (
    snapshot !== undefined &&
    options.expectedRevision !== undefined &&
    snapshot.revision !== options.expectedRevision
  ) {
    throw new TypeError(
      `Permission revision is stale: expected ${options.expectedRevision}, received ${snapshot.revision}`,
    );
  }
  const deniedTargets = requests
    .filter((request) => !(snapshot?.can(request.action, request.target) ?? false))
    .map((request) => snapshotTarget(request.target));
  return Object.freeze({
    allowed: deniedTargets.length === 0,
    snapshotRevision: snapshot?.revision ?? 'unavailable',
    deniedTargets: Object.freeze(deniedTargets),
  });
}

/** Creates an atomic snapshot store used by UI selectors and commit guards. */
export function createPermissionStore(): PermissionStore {
  let current: PermissionSnapshot | undefined;
  const listeners = new Set<(snapshot: PermissionSnapshot | undefined) => void>();
  return Object.freeze({
    getSnapshot: (): PermissionSnapshot | undefined => current,
    replace(snapshot: PermissionSnapshot): void {
      current = snapshot;
      const notificationTargets = new Set(listeners);
      for (const listener of notificationTargets) listener(snapshot);
    },
    clear(): void {
      current = undefined;
      const notificationTargets = new Set(listeners);
      for (const listener of notificationTargets) listener(undefined);
    },
    can(action: PermissionAction, target: PermissionTarget): boolean {
      return current?.can(action, target) ?? false;
    },
    subscribe(listener: (snapshot: PermissionSnapshot | undefined) => void): () => void {
      listeners.add(listener);
      let active = true;
      return (): void => {
        if (!active) return;
        active = false;
        listeners.delete(listener);
      };
    },
  });
}
