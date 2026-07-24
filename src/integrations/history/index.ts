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
}

export interface DocumentDiffOptions {
  readonly signal: AbortSignal;
  readonly maximumUsedCells?: number;
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
