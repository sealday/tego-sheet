import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type { SpreadsheetDocument } from '../../document';
import type { PermissionSnapshot, PermissionTarget } from '../permission';

export type AiContextInclude = 'values' | 'formulas' | 'formats' | 'headers' | 'template-bindings';

export type AiRedactionRule =
  | { readonly kind: 'mask-strings'; readonly replacement: string }
  | { readonly kind: 'omit-formulas' };

export interface AiContextRange {
  readonly sheetId: string;
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
}

export interface ProjectAiContextOptions {
  readonly documentRevision: string;
  readonly ranges: readonly AiContextRange[];
  readonly include: readonly AiContextInclude[];
  readonly redactions: readonly AiRedactionRule[];
  readonly maximumCells?: number;
  readonly maximumBytes?: number;
}

export interface SanitizedDocumentContext {
  readonly schemaVersion: 1;
  readonly documentRevision: string;
  readonly sheets: readonly {
    readonly id: string;
    readonly name: string;
    readonly cells: readonly {
      readonly row: number;
      readonly column: number;
      readonly value?: string | number | boolean | null;
      readonly formula?: string;
    }[];
  }[];
  readonly omittedCellCount: number;
}

export interface AiRequest {
  readonly instruction: string;
  readonly locale: string;
  readonly allowedCommandTypes: readonly WorkbookCommand['type'][];
}

export interface AiCommandProposal {
  readonly id: string;
  readonly summary: string;
  readonly assumptions: readonly string[];
  readonly commands: readonly WorkbookCommand[];
}

export interface AiCommandPort {
  propose(
    request: AiRequest,
    context: SanitizedDocumentContext,
    signal: AbortSignal,
  ): Promise<unknown>;
}

export interface AiTransactionPreview {
  readonly status: 'ready' | 'noop';
  readonly diagnostics: readonly { readonly severity?: string }[];
}

export interface AiProposalSession<ApplyResult = unknown> {
  readonly proposal: AiCommandProposal;
  readonly preview: AiTransactionPreview;
  accept(): ApplyResult;
  reject(): void;
}

export interface CreateAiProposalSessionOptions<ApplyResult> {
  readonly documentId: string;
  readonly documentRevision: string;
  readonly permissionSnapshot: PermissionSnapshot | undefined;
  readonly signal: AbortSignal;
  readonly request: AiRequest;
  readonly context: SanitizedDocumentContext;
  readonly adapter: AiCommandPort;
  readonly dryRun: (commands: readonly WorkbookCommand[]) => AiTransactionPreview;
  readonly apply: (commands: readonly WorkbookCommand[]) => ApplyResult;
  readonly getCurrentRevision: () => string;
  readonly getPermissionSnapshot: () => PermissionSnapshot | undefined;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;
const forbiddenCommandTypes = new Set<string>(['undo', 'redo']);

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function rangeTarget(
  sheetId: string,
  start: { readonly row: number; readonly column: number },
  end = start,
): PermissionTarget {
  return {
    type: 'range',
    range: { sheetId, start, end },
  };
}

function commandPermissionTargets(
  command: WorkbookCommand,
  documentId: string,
): readonly PermissionTarget[] {
  switch (command.type) {
    case 'set-cell-text':
    case 'set-cell-input':
      return [rangeTarget(command.address.sheet, command.address)];
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
      return [
        rangeTarget(
          command.selection.sheet,
          command.selection.range.start,
          command.selection.range.end,
        ),
      ];
    case 'paint-format':
      return [
        rangeTarget(command.target.sheet, command.target.range.start, command.target.range.end),
      ];
    case 'paste-internal':
    case 'paste-external':
      return [
        rangeTarget(command.target.sheet, command.target.range.start, command.target.range.end),
      ];
    case 'autofill':
      return [
        rangeTarget(command.target.sheet, command.target.range.start, command.target.range.end),
      ];
    case 'set-sheet-object':
      return [
        {
          type: 'object',
          sheetId: command.sheet,
          objectId: command.object.id,
        },
      ];
    case 'remove-sheet-object':
      return [{ type: 'object', sheetId: command.sheet, objectId: command.objectId }];
    case 'add-sheet':
      return [{ type: 'document', documentId }];
    case 'undo':
    case 'redo':
      return [{ type: 'document', documentId }];
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
      return [{ type: 'sheet', sheetId: command.sheet }];
  }
}

function inRange(row: number, column: number, range: AiContextRange): boolean {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startColumn = Math.min(range.start.column, range.end.column);
  const endColumn = Math.max(range.start.column, range.end.column);
  return row >= startRow && row <= endRow && column >= startColumn && column <= endColumn;
}

/** Projects only explicitly selected sparse cells into a bounded, credential-free context. */
export function projectAiContext(
  document: SpreadsheetDocument,
  options: ProjectAiContextOptions,
): SanitizedDocumentContext {
  const maximumCells = options.maximumCells ?? 10_000;
  const maximumBytes = options.maximumBytes ?? 1024 * 1024;
  if (!Number.isSafeInteger(maximumCells) || maximumCells < 0 || maximumCells > 100_000) {
    throw new RangeError('AI context maximumCells must be from 0 through 100000');
  }
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 1 || maximumBytes > 8 * 1024 * 1024) {
    throw new RangeError('AI context maximumBytes is invalid');
  }
  const include = new Set(options.include);
  const omitFormulas = options.redactions.some(({ kind }) => kind === 'omit-formulas');
  const mask = options.redactions.find(
    (rule): rule is Extract<AiRedactionRule, { readonly kind: 'mask-strings' }> =>
      rule.kind === 'mask-strings',
  );
  let selectedCount = 0;
  let omittedCellCount = 0;
  const sheets = document.workbook.sheets.flatMap((sheet) => {
    const ranges = options.ranges.filter(({ sheetId }) => sheetId === sheet.id);
    const cells = sheet.cells.flatMap(({ row, column, cell }) => {
      if (!ranges.some((range) => inRange(row, column, range))) {
        omittedCellCount += 1;
        return [];
      }
      selectedCount += 1;
      if (selectedCount > maximumCells) {
        throw new RangeError(`AI context cell limit exceeded: ${maximumCells}`);
      }
      const input = cell.input;
      const output: {
        row: number;
        column: number;
        value?: string | number | boolean | null;
        formula?: string;
      } = { row, column };
      if (input.type === 'formula') {
        if (include.has('formulas') && !omitFormulas) output.formula = input.source;
      } else if (include.has('values')) {
        if (input.type === 'blank') output.value = null;
        else if (input.type === 'custom') output.value = '[custom value omitted]';
        else {
          output.value =
            typeof input.value === 'string' && mask !== undefined ? mask.replacement : input.value;
        }
      }
      return [Object.freeze(output)];
    });
    return cells.length === 0
      ? []
      : [
          Object.freeze({
            id: sheet.id,
            name: sheet.name,
            cells: Object.freeze(cells),
          }),
        ];
  });
  const context: SanitizedDocumentContext = Object.freeze({
    schemaVersion: 1,
    documentRevision: identifier(options.documentRevision, 'AI document revision'),
    sheets: Object.freeze(sheets),
    omittedCellCount,
  });
  const bytes = new TextEncoder().encode(JSON.stringify(context)).byteLength;
  if (bytes > maximumBytes) throw new RangeError(`AI context exceeds ${maximumBytes} bytes`);
  return context;
}

function snapshotProposal(value: unknown, allowed: ReadonlySet<string>): AiCommandProposal {
  let snapshot: unknown;
  try {
    snapshot = JSON.parse(JSON.stringify(value));
  } catch {
    throw new TypeError('AI response is not JSON serializable');
  }
  if (snapshot === null || typeof snapshot !== 'object' || Array.isArray(snapshot)) {
    throw new TypeError('AI response is invalid');
  }
  const proposal = snapshot as Partial<AiCommandProposal>;
  if (
    typeof proposal.id !== 'string' ||
    typeof proposal.summary !== 'string' ||
    !Array.isArray(proposal.assumptions) ||
    !proposal.assumptions.every((entry) => typeof entry === 'string') ||
    !Array.isArray(proposal.commands) ||
    proposal.commands.length > 1_000
  ) {
    throw new TypeError('AI response is invalid');
  }
  for (const command of proposal.commands) {
    const type =
      command !== null && typeof command === 'object' && 'type' in command
        ? (command as { readonly type?: unknown }).type
        : undefined;
    if (typeof type !== 'string' || !allowed.has(type) || forbiddenCommandTypes.has(type)) {
      throw new TypeError(`AI command ${String(type)} is not allowed`);
    }
  }
  return deepFreeze({
    id: identifier(proposal.id, 'AI proposal ID'),
    summary: proposal.summary.slice(0, 2_000),
    assumptions: Object.freeze(proposal.assumptions.map((entry) => entry.slice(0, 2_000))),
    commands: proposal.commands as WorkbookCommand[],
  });
}

function deepFreeze<Value>(value: Value): Value {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

/** Creates a dry-run proposal session that cannot apply without a fresh explicit accept call. */
export async function createAiProposalSession<ApplyResult>(
  options: CreateAiProposalSessionOptions<ApplyResult>,
): Promise<AiProposalSession<ApplyResult>> {
  if (options.signal.aborted) throw new TypeError('AI request was cancelled');
  const initialPermissions = options.permissionSnapshot;
  if (
    initialPermissions === undefined ||
    !initialPermissions.can('ai:propose', {
      type: 'document',
      documentId: options.documentId,
    })
  ) {
    throw new TypeError('AI context permission denied');
  }
  if (options.request.instruction.length > 20_000) {
    throw new RangeError('AI instruction limit is 20000 characters');
  }
  const allowed = new Set<string>(options.request.allowedCommandTypes);
  const response = await options.adapter.propose(options.request, options.context, options.signal);
  if (options.signal.aborted) throw new TypeError('AI request was cancelled');
  const proposal = snapshotProposal(response, allowed);
  const preview = Object.freeze(options.dryRun(proposal.commands));
  if (preview.diagnostics.some(({ severity }) => severity === 'error')) {
    throw new TypeError('AI proposal dry-run produced errors');
  }
  const permissionRevision = initialPermissions.revision;
  let settled = false;
  return Object.freeze({
    proposal,
    preview,
    accept(): ApplyResult {
      if (settled) throw new TypeError('AI proposal session is already settled');
      if (options.getCurrentRevision() !== options.documentRevision) {
        throw new TypeError('AI proposal preview is stale');
      }
      const permissions = options.getPermissionSnapshot();
      if (
        permissions?.revision !== permissionRevision ||
        !permissions.can('ai:apply', {
          type: 'document',
          documentId: options.documentId,
        })
      ) {
        throw new TypeError('AI proposal permission is stale or denied');
      }
      const deniedTarget = proposal.commands
        .flatMap((command) => commandPermissionTargets(command, options.documentId))
        .find((target) => !permissions.can('ai:apply', target));
      if (deniedTarget !== undefined) {
        throw new TypeError('AI proposal target permission is denied');
      }
      settled = true;
      return options.apply(proposal.commands);
    },
    reject(): void {
      settled = true;
    },
  });
}
