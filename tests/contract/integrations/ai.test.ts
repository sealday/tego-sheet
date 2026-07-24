import { describe, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument } from '../../../src/document';
import { createAiProposalSession, projectAiContext } from '../../../src/integrations/ai';
import { createPermissionSnapshot } from '../../../src/integrations/permission';

function document() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'document-1',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [
            { row: 0, column: 0, cell: { input: { type: 'string', value: 'secret' } } },
            { row: 0, column: 1, cell: { input: { type: 'number', value: 2 } } },
            { row: 5, column: 5, cell: { input: { type: 'string', value: 'hidden' } } },
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
    extensions: { 'host.secret': { apiKey: 'never-send' } },
  });
  if (!parsed.ok) throw new Error('AI fixture failed');
  return parsed.document;
}

const permissions = createPermissionSnapshot({
  revision: 'permission-1',
  actorId: 'actor-1',
  grants: [
    {
      action: 'ai:propose',
      target: { type: 'document', documentId: 'document-1' },
    },
    {
      action: 'ai:apply',
      target: { type: 'document', documentId: 'document-1' },
    },
    {
      action: 'ai:apply',
      target: {
        type: 'range',
        range: {
          sheetId: 'sheet-1',
          start: { row: 0, column: 0 },
          end: { row: 0, column: 0 },
        },
      },
    },
  ],
});

describe('AI command integration contract', () => {
  it('projects only selected cells and applies explicit redaction', () => {
    const context = projectAiContext(document(), {
      documentRevision: 'revision-1',
      ranges: [
        {
          sheetId: 'sheet-1',
          start: { row: 0, column: 0 },
          end: { row: 0, column: 1 },
        },
      ],
      include: ['values'],
      redactions: [{ kind: 'mask-strings', replacement: '[redacted]' }],
    });

    expect(context.sheets[0]?.cells).toEqual([
      { row: 0, column: 0, value: '[redacted]' },
      { row: 0, column: 1, value: 2 },
    ]);
    expect(JSON.stringify(context)).not.toContain('hidden');
    expect(JSON.stringify(context)).not.toContain('apiKey');
  });

  it('rejects unknown or disallowed proposal commands before dry-run', async () => {
    const dryRun = vi.fn();
    await expect(
      createAiProposalSession({
        documentId: 'document-1',
        documentRevision: 'revision-1',
        permissionSnapshot: permissions,
        signal: new AbortController().signal,
        request: {
          instruction: 'print it',
          locale: 'en-US',
          allowedCommandTypes: ['set-cell-input'],
        },
        context: projectAiContext(document(), {
          documentRevision: 'revision-1',
          ranges: [],
          include: [],
          redactions: [],
        }),
        adapter: {
          propose: async () => ({
            id: 'proposal-1',
            summary: 'Unsafe',
            assumptions: [],
            commands: [{ type: 'undo' }],
          }),
        },
        dryRun,
        apply: vi.fn(),
        getCurrentRevision: () => 'revision-1',
        getPermissionSnapshot: () => permissions,
      }),
    ).rejects.toThrow(/not allowed/u);
    expect(dryRun).not.toHaveBeenCalled();
  });

  it('requires explicit accept and expires on document revision change', async () => {
    const apply = vi.fn(() => ({ status: 'committed' }));
    let revision = 'revision-1';
    const session = await createAiProposalSession({
      documentId: 'document-1',
      documentRevision: revision,
      permissionSnapshot: permissions,
      signal: new AbortController().signal,
      request: {
        instruction: 'set A1',
        locale: 'en-US',
        allowedCommandTypes: ['set-cell-input'],
      },
      context: projectAiContext(document(), {
        documentRevision: revision,
        ranges: [],
        include: [],
        redactions: [],
      }),
      adapter: {
        propose: async () => ({
          id: 'proposal-1',
          summary: 'Set A1',
          assumptions: [],
          commands: [
            {
              type: 'set-cell-input',
              address: { sheet: 'sheet-1', row: 0, column: 0 },
              input: { type: 'string', value: 'done' },
            },
          ],
        }),
      },
      dryRun: vi.fn(() => ({ status: 'ready' as const, diagnostics: [] })),
      apply,
      getCurrentRevision: () => revision,
      getPermissionSnapshot: () => permissions,
    });

    expect(apply).not.toHaveBeenCalled();
    revision = 'revision-2';
    expect(() => session.accept()).toThrow(/stale/u);
    expect(apply).not.toHaveBeenCalled();
  });

  it('deeply snapshots commands and checks every target before applying', async () => {
    const source = {
      id: 'proposal-1',
      summary: 'Set A1',
      assumptions: [],
      commands: [
        {
          type: 'set-cell-input' as const,
          address: { sheet: 'sheet-1', row: 0, column: 0 },
          input: { type: 'string' as const, value: 'safe' },
        },
      ],
    };
    const apply = vi.fn();
    let currentPermissions = permissions;
    const session = await createAiProposalSession({
      documentId: 'document-1',
      documentRevision: 'revision-1',
      permissionSnapshot: permissions,
      signal: new AbortController().signal,
      request: {
        instruction: 'set A1',
        locale: 'en-US',
        allowedCommandTypes: ['set-cell-input'],
      },
      context: projectAiContext(document(), {
        documentRevision: 'revision-1',
        ranges: [],
        include: [],
        redactions: [],
      }),
      adapter: { propose: async () => source },
      dryRun: vi.fn(() => ({ status: 'ready' as const, diagnostics: [] })),
      apply,
      getCurrentRevision: () => 'revision-1',
      getPermissionSnapshot: () => currentPermissions,
    });

    source.commands[0]!.input.value = 'tampered';
    expect(session.proposal.commands[0]).toMatchObject({
      input: { value: 'safe' },
    });
    expect(Object.isFrozen((session.proposal.commands[0] as { input: object }).input)).toBe(true);

    currentPermissions = createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [
        {
          action: 'ai:apply',
          target: { type: 'document', documentId: 'document-1' },
        },
      ],
    });
    expect(() => session.accept()).toThrow(/target permission/u);
    expect(apply).not.toHaveBeenCalled();
  });
});
