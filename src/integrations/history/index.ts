import { parseSpreadsheetDocument, type SpreadsheetDocument, type Sheet } from '../../document';
import type { PermissionSnapshot } from '../permission';

export interface DocumentVersionSnapshot {
  readonly id: string;
  readonly document: SpreadsheetDocument;
}

export interface SheetDiff {
  readonly sheetId: string;
  readonly change: 'added' | 'removed' | 'modified';
  readonly nameChanged: boolean;
  readonly cellsChanged: number;
  readonly formulasChanged: number;
  readonly structuralChanges: number;
}

export interface DocumentDiff {
  readonly fromVersion: string;
  readonly toVersion: string;
  readonly sheets: readonly SheetDiff[];
  readonly summary: {
    readonly cellsChanged: number;
    readonly formulasChanged: number;
    readonly structuralChanges: number;
    readonly templatesChanged: number;
    readonly printProfilesChanged: number;
  };
  readonly degraded?: 'structural-summary';
}

export interface DocumentDiffOptions {
  readonly signal: AbortSignal;
  readonly maximumUsedCells?: number;
}

export interface AsyncDocumentDiffOptions extends DocumentDiffOptions {
  readonly onCellBudgetExceeded?: 'throw' | 'structural-summary';
}

export interface HistoryVersionInfo {
  readonly id: string;
  readonly revision: string;
  readonly createdAt: number;
  readonly label?: string;
}

export interface HistoryAdapter {
  list(documentId: string, signal: AbortSignal): Promise<unknown>;
  load(documentId: string, versionId: string, signal: AbortSignal): Promise<unknown>;
  checkpoint(
    request: {
      readonly documentId: string;
      readonly revision: string;
      readonly document: SpreadsheetDocument;
    },
    signal: AbortSignal,
  ): Promise<unknown>;
  commitRestore(proposal: RestoreVersionProposal, signal: AbortSignal): Promise<unknown>;
}

export interface HistoryPreview {
  readonly versionId: string;
  readonly revision: string;
  readonly document: SpreadsheetDocument;
  readonly readOnly: true;
}

export interface HistoryCheckpointResult {
  readonly versionId: string;
  readonly revision: string;
}

export interface RestoreVersionProposal {
  readonly type: 'restore-version';
  readonly sourceVersionId: string;
  readonly expectedCurrentRevision: string;
  readonly replacement: SpreadsheetDocument;
}

export interface CreateRestoreVersionProposalOptions {
  readonly sourceVersionId: string;
  readonly expectedCurrentRevision: string;
  readonly currentRevision: string;
  readonly replacement: SpreadsheetDocument;
  readonly documentId: string;
  readonly permissions: PermissionSnapshot | undefined;
}

const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/;

function identifier(value: string, label: string): string {
  if (!identifierPattern.test(value)) throw new TypeError(`${label} is invalid`);
  return value;
}

function canonical(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
  const record = value as Readonly<Record<string, unknown>>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${canonical(record[key])}`)
    .join(',')}}`;
}

function cellMap(sheet: Sheet): ReadonlyMap<string, Sheet['cells'][number]['cell']> {
  return new Map(sheet.cells.map(({ row, column, cell }) => [`${row}:${column}`, cell]));
}

function structuralFingerprint(sheet: Sheet): string {
  return canonical({
    merges: sheet.merges,
    rows: sheet.rows,
    columns: sheet.columns,
    groups: sheet.groups,
    freeze: sheet.freeze,
    filter: sheet.filter,
    visibility: sheet.visibility,
    tables: sheet.tables,
    objects: sheet.objects,
  });
}

function countChangedEntries(
  left: readonly { readonly id: string }[],
  right: readonly { readonly id: string }[],
): number {
  const leftById = new Map(left.map((entry) => [entry.id, canonical(entry)]));
  const rightById = new Map(right.map((entry) => [entry.id, canonical(entry)]));
  const ids = new Set([...leftById.keys(), ...rightById.keys()]);
  let changed = 0;
  for (const id of ids) if (leftById.get(id) !== rightById.get(id)) changed += 1;
  return changed;
}

function snapshotJson<Value>(value: Value, label: string): Value {
  try {
    return JSON.parse(JSON.stringify(value)) as Value;
  } catch {
    throw new TypeError(`${label} must be JSON serializable`);
  }
}

/** Computes a stable-ID semantic history diff within a hard used-cell budget. */
export function diffDocumentVersions(
  from: DocumentVersionSnapshot,
  to: DocumentVersionSnapshot,
  options: DocumentDiffOptions,
): DocumentDiff {
  if (options.signal.aborted) throw new TypeError('History diff was cancelled');
  const maximumUsedCells = options.maximumUsedCells ?? 1_000_000;
  if (!Number.isSafeInteger(maximumUsedCells) || maximumUsedCells < 1) {
    throw new RangeError('History diff maximumUsedCells must be a positive safe integer');
  }
  const usedCells =
    from.document.workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0) +
    to.document.workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0);
  if (usedCells > maximumUsedCells) {
    throw new RangeError(`History diff used-cell limit exceeded: ${maximumUsedCells}`);
  }
  const fromSheets = new Map(from.document.workbook.sheets.map((sheet) => [sheet.id, sheet]));
  const toSheets = new Map(to.document.workbook.sheets.map((sheet) => [sheet.id, sheet]));
  const sheetIds = [...new Set([...fromSheets.keys(), ...toSheets.keys()])].sort();
  const sheets: SheetDiff[] = [];
  for (const sheetId of sheetIds) {
    if (options.signal.aborted) throw new TypeError('History diff was cancelled');
    const left = fromSheets.get(sheetId);
    const right = toSheets.get(sheetId);
    if (left === undefined || right === undefined) {
      const existing = left ?? right!;
      sheets.push(
        Object.freeze({
          sheetId,
          change: left === undefined ? 'added' : 'removed',
          nameChanged: false,
          cellsChanged: existing.cells.filter(({ cell }) => cell.input.type !== 'formula').length,
          formulasChanged: existing.cells.filter(({ cell }) => cell.input.type === 'formula')
            .length,
          structuralChanges: 1,
        }),
      );
      continue;
    }
    const leftCells = cellMap(left);
    const rightCells = cellMap(right);
    const addresses = new Set([...leftCells.keys(), ...rightCells.keys()]);
    let cellsChanged = 0;
    let formulasChanged = 0;
    for (const address of addresses) {
      const before = leftCells.get(address);
      const after = rightCells.get(address);
      if (canonical(before) === canonical(after)) continue;
      if (before?.input.type === 'formula' || after?.input.type === 'formula') formulasChanged += 1;
      else cellsChanged += 1;
    }
    const structuralChanges = structuralFingerprint(left) === structuralFingerprint(right) ? 0 : 1;
    if (
      left.name !== right.name ||
      cellsChanged > 0 ||
      formulasChanged > 0 ||
      structuralChanges > 0
    ) {
      sheets.push(
        Object.freeze({
          sheetId,
          change: 'modified',
          nameChanged: left.name !== right.name,
          cellsChanged,
          formulasChanged,
          structuralChanges,
        }),
      );
    }
  }
  const templatesChanged = countChangedEntries(from.document.templates, to.document.templates);
  const fromPrintProfiles = from.document.templates.flatMap((template) =>
    template.printProfiles.map((profile) => ({
      id: `${template.id}:${profile.id}`,
      value: profile,
    })),
  );
  const toPrintProfiles = to.document.templates.flatMap((template) =>
    template.printProfiles.map((profile) => ({
      id: `${template.id}:${profile.id}`,
      value: profile,
    })),
  );
  const printProfilesChanged = countChangedEntries(fromPrintProfiles, toPrintProfiles);
  const summary = Object.freeze({
    cellsChanged: sheets.reduce((sum, sheet) => sum + sheet.cellsChanged, 0),
    formulasChanged: sheets.reduce((sum, sheet) => sum + sheet.formulasChanged, 0),
    structuralChanges: sheets.reduce((sum, sheet) => sum + sheet.structuralChanges, 0),
    templatesChanged,
    printProfilesChanged,
  });
  return Object.freeze({
    fromVersion: identifier(from.id, 'History fromVersion'),
    toVersion: identifier(to.id, 'History toVersion'),
    sheets: Object.freeze(sheets),
    summary,
  });
}

/** Computes a cancellable diff and can degrade to a bounded structural-only summary. */
export async function diffDocumentVersionsAsync(
  from: DocumentVersionSnapshot,
  to: DocumentVersionSnapshot,
  options: AsyncDocumentDiffOptions,
): Promise<DocumentDiff> {
  await Promise.resolve();
  if (options.signal.aborted) throw new TypeError('History diff was cancelled');
  try {
    return diffDocumentVersions(from, to, options);
  } catch (error) {
    if (!(error instanceof RangeError) || options.onCellBudgetExceeded !== 'structural-summary') {
      throw error;
    }
    const withoutCells = (snapshot: DocumentVersionSnapshot): DocumentVersionSnapshot => ({
      ...snapshot,
      document: {
        ...snapshot.document,
        workbook: {
          ...snapshot.document.workbook,
          sheets: snapshot.document.workbook.sheets.map((sheet) => ({ ...sheet, cells: [] })),
        },
      },
    });
    const result = diffDocumentVersions(withoutCells(from), withoutCells(to), {
      signal: options.signal,
      maximumUsedCells: 1,
    });
    return Object.freeze({ ...result, degraded: 'structural-summary' });
  }
}

/** Lists bounded immutable version metadata supplied by the host. */
export async function listDocumentVersions(
  adapter: Pick<HistoryAdapter, 'list'>,
  documentId: string,
  signal: AbortSignal,
): Promise<readonly HistoryVersionInfo[]> {
  const id = identifier(documentId, 'History documentId');
  const value = snapshotJson(await adapter.list(id, signal), 'History version list');
  if (signal.aborted) throw new TypeError('History list was cancelled');
  if (!Array.isArray(value) || value.length > 10_000) {
    throw new RangeError('History version list limit is 10000');
  }
  return Object.freeze(
    value.map((entry) => {
      if (
        entry === null ||
        typeof entry !== 'object' ||
        Array.isArray(entry) ||
        typeof (entry as HistoryVersionInfo).createdAt !== 'number'
      ) {
        throw new TypeError('History version metadata is invalid');
      }
      const item = entry as HistoryVersionInfo;
      return Object.freeze({
        id: identifier(item.id, 'History versionId'),
        revision: identifier(item.revision, 'History revision'),
        createdAt: item.createdAt,
        ...(typeof item.label === 'string' ? { label: item.label.slice(0, 1_000) } : {}),
      });
    }),
  );
}

/** Loads one isolated read-only version preview after full schema validation. */
export async function loadHistoryPreview(
  adapter: Pick<HistoryAdapter, 'load'>,
  documentId: string,
  versionId: string,
  signal: AbortSignal,
): Promise<HistoryPreview> {
  const expectedDocumentId = identifier(documentId, 'History documentId');
  const expectedVersionId = identifier(versionId, 'History versionId');
  const value = snapshotJson(
    await adapter.load(expectedDocumentId, expectedVersionId, signal),
    'History version',
  ) as {
    readonly id?: unknown;
    readonly documentId?: unknown;
    readonly revision?: unknown;
    readonly document?: unknown;
  };
  if (signal.aborted) throw new TypeError('History load was cancelled');
  if (
    value.id !== expectedVersionId ||
    value.documentId !== expectedDocumentId ||
    typeof value.revision !== 'string'
  ) {
    throw new TypeError('History version identity is invalid');
  }
  const parsed = parseSpreadsheetDocument(value.document);
  if (!parsed.ok || parsed.document.id !== expectedDocumentId) {
    throw new TypeError('History version document is invalid');
  }
  return Object.freeze({
    versionId: expectedVersionId,
    revision: identifier(value.revision, 'History revision'),
    document: parsed.document,
    readOnly: true,
  });
}

/** Creates one immutable host checkpoint request and validates its acknowledgement. */
export async function checkpointDocumentVersion(
  adapter: Pick<HistoryAdapter, 'checkpoint'>,
  request: {
    readonly documentId: string;
    readonly revision: string;
    readonly document: SpreadsheetDocument;
  },
  signal: AbortSignal,
): Promise<HistoryCheckpointResult> {
  const documentId = identifier(request.documentId, 'History documentId');
  const revision = identifier(request.revision, 'History revision');
  const parsed = parseSpreadsheetDocument(snapshotJson(request.document, 'History checkpoint'));
  if (!parsed.ok || parsed.document.id !== documentId) {
    throw new TypeError('History checkpoint document is invalid');
  }
  const acknowledgement = snapshotJson(
    await adapter.checkpoint(
      Object.freeze({ documentId, revision, document: parsed.document }),
      signal,
    ),
    'History checkpoint acknowledgement',
  ) as Partial<HistoryCheckpointResult>;
  if (signal.aborted) throw new TypeError('History checkpoint was cancelled');
  return Object.freeze({
    versionId: identifier(acknowledgement.versionId ?? '', 'History checkpoint versionId'),
    revision: identifier(acknowledgement.revision ?? '', 'History checkpoint revision'),
  });
}

/** Creates an explicit restore proposal which must later become a new transaction/checkpoint. */
export function createRestoreVersionProposal(
  options: CreateRestoreVersionProposalOptions,
): RestoreVersionProposal {
  if (options.expectedCurrentRevision !== options.currentRevision) {
    throw new TypeError('History restore proposal is stale');
  }
  if (
    !(
      options.permissions?.can('history:restore', {
        type: 'document',
        documentId: options.documentId,
      }) ?? false
    )
  ) {
    throw new TypeError('History restore permission denied');
  }
  const parsed = parseSpreadsheetDocument(JSON.parse(JSON.stringify(options.replacement)));
  if (!parsed.ok) throw new TypeError('History replacement document is invalid');
  return Object.freeze({
    type: 'restore-version',
    sourceVersionId: identifier(options.sourceVersionId, 'History sourceVersionId'),
    expectedCurrentRevision: identifier(
      options.expectedCurrentRevision,
      'History current revision',
    ),
    replacement: parsed.document,
  });
}

/** Commits one validated restore proposal as a single host-owned checkpoint operation. */
export async function restoreDocumentVersion(
  adapter: Pick<HistoryAdapter, 'commitRestore'>,
  options: CreateRestoreVersionProposalOptions,
  signal: AbortSignal,
): Promise<HistoryCheckpointResult> {
  if (signal.aborted) throw new TypeError('History restore was cancelled');
  const proposal = createRestoreVersionProposal(options);
  const acknowledgement = snapshotJson(
    await adapter.commitRestore(proposal, signal),
    'History restore acknowledgement',
  ) as Partial<HistoryCheckpointResult>;
  if (signal.aborted) throw new TypeError('History restore was cancelled');
  return Object.freeze({
    versionId: identifier(acknowledgement.versionId ?? '', 'History restore versionId'),
    revision: identifier(acknowledgement.revision ?? '', 'History restore revision'),
  });
}
