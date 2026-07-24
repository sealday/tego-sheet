import { describe, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument } from '../../../src/document';
import {
  checkpointDocumentVersion,
  createRestoreVersionProposal,
  diffDocumentVersions,
  diffDocumentVersionsAsync,
  loadHistoryPreview,
  restoreDocumentVersion,
} from '../../../src/integrations/history';
import { createPermissionSnapshot } from '../../../src/integrations/permission';

function document(name: string, value: number) {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'document-1',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name,
          cells: [
            { row: 0, column: 0, cell: { input: { type: 'number', value } } },
            { row: 0, column: 1, cell: { input: { type: 'formula', source: '=A1*2' } } },
          ],
          merges: [],
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error('history fixture failed');
  return parsed.document;
}

function withPrintProfile(base: ReturnType<typeof document>) {
  const parsed = parseSpreadsheetDocument({
    ...base,
    templates: [
      {
        id: 'template-1',
        name: 'Invoice',
        bindings: [],
        printProfiles: [
          {
            id: 'profile-1',
            name: 'Default',
            targets: [],
            page: {
              size: { kind: 'named', name: 'A4' },
              orientation: 'portrait',
              margins: { top: 1, right: 1, bottom: 1, left: 1 },
              scale: { mode: 'actual-size' },
            },
            manualBreaks: [],
            showGridlines: false,
            showHeadings: false,
          },
        ],
      },
    ],
  });
  if (!parsed.ok) throw new Error('history print fixture failed');
  return parsed.document;
}

describe('version history integration contract', () => {
  it('diffs by stable identity so a sheet rename is not deletion plus creation', () => {
    const result = diffDocumentVersions(
      { id: 'version-1', document: document('Old name', 1) },
      { id: 'version-2', document: document('New name', 2) },
      { signal: new AbortController().signal },
    );

    expect(result.summary).toEqual({
      cellsChanged: 1,
      formulasChanged: 0,
      structuralChanges: 0,
      templatesChanged: 0,
      printProfilesChanged: 0,
    });
    expect(result.sheets).toEqual([
      expect.objectContaining({
        sheetId: 'sheet-1',
        change: 'modified',
        nameChanged: true,
      }),
    ]);
  });

  it('refuses stale or unauthorized restore proposals', () => {
    const denied = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [],
    });
    expect(() =>
      createRestoreVersionProposal({
        sourceVersionId: 'version-1',
        expectedCurrentRevision: 'revision-1',
        currentRevision: 'revision-2',
        replacement: document('Old', 1),
        documentId: 'document-1',
        permissions: denied,
      }),
    ).toThrow(/stale/u);
    expect(() =>
      createRestoreVersionProposal({
        sourceVersionId: 'version-1',
        expectedCurrentRevision: 'revision-2',
        currentRevision: 'revision-2',
        replacement: document('Old', 1),
        documentId: 'document-1',
        permissions: denied,
      }),
    ).toThrow(/denied/u);
  });

  it('creates a proposal for a new restore transaction without mutating either snapshot', () => {
    const replacement = document('Old', 1);
    const permissions = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [
        {
          action: 'history:restore',
          target: { type: 'document', documentId: 'document-1' },
        },
      ],
    });

    const proposal = createRestoreVersionProposal({
      sourceVersionId: 'version-1',
      expectedCurrentRevision: 'revision-2',
      currentRevision: 'revision-2',
      replacement,
      documentId: 'document-1',
      permissions,
    });

    expect(proposal).toEqual({
      type: 'restore-version',
      sourceVersionId: 'version-1',
      expectedCurrentRevision: 'revision-2',
      replacement,
    });
    expect(Object.isFrozen(proposal)).toBe(true);
  });

  it('counts print profiles added with a new template', () => {
    const result = diffDocumentVersions(
      { id: 'version-1', document: document('Sheet', 1) },
      { id: 'version-2', document: withPrintProfile(document('Sheet', 1)) },
      { signal: new AbortController().signal },
    );

    expect(result.summary.templatesChanged).toBe(1);
    expect(result.summary.printProfilesChanged).toBe(1);
  });

  it('loads isolated readonly previews and checkpoints through host adapters', async () => {
    const source = document('Old', 1);
    const preview = await loadHistoryPreview(
      {
        load: async () => ({
          id: 'version-1',
          documentId: source.id,
          revision: 'revision-1',
          document: source,
        }),
      },
      source.id,
      'version-1',
      new AbortController().signal,
    );
    expect(preview.readOnly).toBe(true);
    expect(Object.isFrozen(preview.document)).toBe(true);

    const checkpoint = await checkpointDocumentVersion(
      {
        checkpoint: async (request) => ({
          versionId: 'version-2',
          revision: request.revision,
        }),
      },
      {
        documentId: source.id,
        revision: 'revision-2',
        document: source,
      },
      new AbortController().signal,
    );
    expect(checkpoint).toEqual({ versionId: 'version-2', revision: 'revision-2' });
  });

  it('provides cancellable async structural summaries when cell budgets are exceeded', async () => {
    const result = await diffDocumentVersionsAsync(
      { id: 'version-1', document: document('Old', 1) },
      { id: 'version-2', document: document('New', 2) },
      {
        signal: new AbortController().signal,
        maximumUsedCells: 1,
        onCellBudgetExceeded: 'structural-summary',
      },
    );
    expect(result.degraded).toBe('structural-summary');
    expect(result.sheets[0]).toMatchObject({ nameChanged: true });
  });

  it('restores through one host checkpoint operation', async () => {
    const replacement = document('Old', 1);
    const permissions = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [
        {
          action: 'history:restore',
          target: { type: 'document', documentId: replacement.id },
        },
      ],
    });
    const commitRestore = vi.fn(async () => ({
      versionId: 'version-3',
      revision: 'revision-3',
    }));
    const result = await restoreDocumentVersion(
      { commitRestore },
      {
        sourceVersionId: 'version-1',
        expectedCurrentRevision: 'revision-2',
        currentRevision: 'revision-2',
        replacement,
        documentId: replacement.id,
        permissions,
      },
      new AbortController().signal,
    );
    expect(result).toEqual({ versionId: 'version-3', revision: 'revision-3' });
    expect(commitRestore).toHaveBeenCalledTimes(1);
  });
});
