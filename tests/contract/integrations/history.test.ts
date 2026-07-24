import { describe, expect, it } from 'vitest';
import { parseSpreadsheetDocument } from '../../../src/document';
import {
  createRestoreVersionProposal,
  diffDocumentVersions,
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
});
