import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type { SpreadsheetDocument } from '../../document';
import type {
  DocumentCommand,
  DocumentController,
  DocumentTransactionEnvelope,
  DocumentTransactionResult,
} from '../../document-controller';
import type {
  PermissionAction,
  PermissionSnapshot,
  PermissionStore,
  PermissionTarget,
} from '../permission';

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
  readonly baseRevision?: number;
  readonly document?: SpreadsheetDocument;
}

export interface AiProposalSession<ApplyResult = unknown> {
  readonly proposal: AiCommandProposal;
  readonly preview: AiTransactionPreview;
  accept(): ApplyResult;
  reject(): void;
}

/** Value-free summary shown before or alongside an AI request. */
export interface AiContextSummary {
  readonly sheetCount: number;
  readonly cellCount: number;
  readonly omittedCellCount: number;
  readonly serializedBytes: number;
}

/** Proposal session bound to the public atomic document controller. */
export interface ControllerAiProposalSession extends AiProposalSession<DocumentTransactionResult> {
  readonly contextSummary: AiContextSummary;
}

export interface CreateControllerAiProposalSessionOptions {
  readonly controller: DocumentController;
  readonly permissions: PermissionStore;
  readonly adapter: AiCommandPort;
  readonly signal: AbortSignal;
  readonly request: AiRequest;
  readonly context: Omit<ProjectAiContextOptions, 'documentRevision'>;
  /** Stable caller-owned ID used for the one previewed and committed transaction. */
  readonly transactionId: string;
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

interface CommandPermissionRequest {
  readonly action: PermissionAction;
  readonly target: PermissionTarget;
}

function commandPermissionRequests(
  command: WorkbookCommand,
  documentId: string,
): readonly CommandPermissionRequest[] {
  switch (command.type) {
    case 'set-cell-text':
    case 'set-cell-input':
      return [
        { action: 'range:edit', target: rangeTarget(command.address.sheet, command.address) },
      ];
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
        {
          action: 'range:edit',
          target: rangeTarget(
            command.selection.sheet,
            command.selection.range.start,
            command.selection.range.end,
          ),
        },
      ];
    case 'paint-format':
      return [
        {
          action: 'range:edit',
          target: rangeTarget(
            command.target.sheet,
            command.target.range.start,
            command.target.range.end,
          ),
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId: command.source.sheet },
        },
      ];
    case 'paste-internal': {
      const requests: CommandPermissionRequest[] = [
        {
          action: 'range:edit',
          target: rangeTarget(
            command.target.sheet,
            command.target.range.start,
            command.target.range.end,
          ),
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId: command.source.sheet },
        },
      ];
      if (command.cut) {
        requests.push({
          action: 'range:edit',
          target: rangeTarget(
            command.source.sheet,
            command.source.range.start,
            command.source.range.end,
          ),
        });
      }
      return requests;
    }
    case 'paste-external':
      return [
        {
          action: 'range:edit',
          target: rangeTarget(
            command.target.sheet,
            command.target.range.start,
            command.target.range.end,
          ),
        },
      ];
    case 'autofill':
      return [
        {
          action: 'range:edit',
          target: rangeTarget(
            command.target.sheet,
            command.target.range.start,
            command.target.range.end,
          ),
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId: command.source.sheet },
        },
      ];
    case 'set-sheet-object':
      return [
        {
          action: 'object:edit',
          target: {
            type: 'object',
            sheetId: command.sheet,
            objectId: command.object.id,
          },
        },
      ];
    case 'remove-sheet-object':
      return [
        {
          action: 'object:edit',
          target: { type: 'object', sheetId: command.sheet, objectId: command.objectId },
        },
      ];
    case 'set-chart': {
      const sourceSheets = new Set<string>();
      if (command.chart.categories !== undefined) {
        sourceSheets.add(command.chart.categories.sheetId);
      }
      for (const series of command.chart.series) {
        sourceSheets.add(series.values.sheetId);
      }
      return [
        {
          action: 'object:edit',
          target: { type: 'object', sheetId: command.sheet, objectId: command.chart.id },
        },
        ...Array.from(sourceSheets, (sheetId) => ({
          action: 'sheet:view' as const,
          target: { type: 'sheet' as const, sheetId },
        })),
      ];
    }
    case 'remove-chart':
      return [
        {
          action: 'object:edit',
          target: { type: 'object', sheetId: command.sheet, objectId: command.chartId },
        },
      ];
    case 'set-sparkline':
      return [
        {
          action: 'range:edit',
          target: rangeTarget(command.sparkline.target.sheetId, command.sparkline.target),
        },
        {
          action: 'sheet:view',
          target: { type: 'sheet', sheetId: command.sparkline.source.sheetId },
        },
      ];
    case 'remove-sparkline':
      return [{ action: 'sheet:edit', target: { type: 'sheet', sheetId: command.sheet } }];
    case 'add-sheet':
      return [{ action: 'document:edit', target: { type: 'document', documentId } }];
    case 'undo':
    case 'redo':
      return [{ action: 'document:edit', target: { type: 'document', documentId } }];
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
      return [{ action: 'sheet:edit', target: { type: 'sheet', sheetId: command.sheet } }];
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

/** Summarizes projected context without exposing cell values or formulas. */
export function summarizeAiContext(context: SanitizedDocumentContext): AiContextSummary {
  return Object.freeze({
    sheetCount: context.sheets.length,
    cellCount: context.sheets.reduce((total, sheet) => total + sheet.cells.length, 0),
    omittedCellCount: context.omittedCellCount,
    serializedBytes: new TextEncoder().encode(JSON.stringify(context)).byteLength,
  });
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
      const deniedRequest = proposal.commands
        .flatMap((command) => commandPermissionRequests(command, options.documentId))
        .find(({ action, target }) => !permissions.can(action, target));
      if (deniedRequest !== undefined) {
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

/** Binds sanitized AI proposals to one dry-run and one atomic public controller transaction. */
export async function createControllerAiProposalSession(
  options: CreateControllerAiProposalSessionOptions,
): Promise<ControllerAiProposalSession> {
  const snapshot = options.controller.getSnapshot();
  const documentRevision = `revision-${snapshot.revision}`;
  const context = projectAiContext(snapshot.document, {
    ...options.context,
    documentRevision,
  });
  const transactionId = identifier(options.transactionId, 'AI transaction ID');
  let transaction: DocumentTransactionEnvelope | undefined;
  const session = await createAiProposalSession({
    documentId: snapshot.document.id,
    documentRevision,
    permissionSnapshot: options.permissions.getSnapshot(),
    signal: options.signal,
    request: options.request,
    context,
    adapter: options.adapter,
    dryRun(commands): AiTransactionPreview {
      transaction = Object.freeze({
        schemaVersion: 1,
        id: transactionId,
        baseRevision: snapshot.revision,
        commands: Object.freeze(
          commands.map((command, index) =>
            Object.freeze({
              schemaVersion: 1,
              id: `${transactionId}:command-${index + 1}`,
              command: command as DocumentCommand,
            }),
          ),
        ),
      });
      const preview = options.controller.dryRun(transaction, { source: 'ref' });
      if (preview.status === 'rejected') {
        return Object.freeze({
          status: 'noop',
          diagnostics: Object.freeze([
            Object.freeze({ severity: 'error', code: preview.code, message: preview.message }),
          ]),
        });
      }
      return Object.freeze({
        status: preview.status,
        diagnostics: Object.freeze([]),
        baseRevision: preview.baseRevision,
        document: preview.document,
      });
    },
    apply(): DocumentTransactionResult {
      if (transaction === undefined) {
        throw new TypeError('AI proposal transaction was not prepared');
      }
      return options.controller.transact(transaction, { source: 'ref' });
    },
    getCurrentRevision: () => `revision-${options.controller.getSnapshot().revision}`,
    getPermissionSnapshot: () => options.permissions.getSnapshot(),
  });
  return Object.freeze({
    ...session,
    contextSummary: summarizeAiContext(context),
  });
}
