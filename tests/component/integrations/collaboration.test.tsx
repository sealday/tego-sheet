import { cleanup, render } from '@testing-library/react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet } from '../../../src';
import {
  createCollaborationSession,
  createPresenceStore,
  createRemoteOperationProcessor,
} from '../../../src/integrations/collaboration';
import { createCanvasHarness } from '../../helpers/canvas-harness';
import { testDocument } from '../../helpers/workbook-builders';

beforeEach(() => {
  const context = createCanvasHarness().canvas.getContext('2d');
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => context);
  vi.stubGlobal(
    'requestAnimationFrame',
    vi.fn(() => 1),
  );
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

it('projects ephemeral remote presence without persisting it in the document', () => {
  const document = testDocument([{ name: 'A' }]);
  const presence = createPresenceStore({ now: () => 0 });
  presence.replace([
    {
      actorId: 'actor-2',
      sheetId: document.workbook.sheets[0]!.id,
      selections: [],
      display: { label: 'Remote user', color: '#ff0000' },
      expiresAt: 100,
    },
  ]);
  const rendered = render(<TegoSheet defaultDocument={document} presenceStore={presence} />);

  expect(rendered.container.querySelector('[data-presence-actor="actor-2"]')).not.toBeNull();
  expect(JSON.stringify(document)).not.toContain('actor-2');
});

it('projects the collaboration connection state as accessible status', async () => {
  const document = testDocument([{ name: 'A' }]);
  const presence = createPresenceStore({ now: () => 0 });
  const collaborationSession = createCollaborationSession({
    presence,
    processor: createRemoteOperationProcessor({
      initialRevision: 'revision-0',
      permissionGate: () => true,
      transactionBoundary: {
        prepare: () => ({ commit: vi.fn(), rollback: vi.fn() }),
      },
    }),
    port: {
      connect: async () => ({
        revision: 'revision-1',
        capabilities: { protocolVersions: [1], collaborativeUndo: true },
      }),
      subscribe: () => () => undefined,
    },
  });
  const rendered = render(
    <TegoSheet
      defaultDocument={document}
      presenceStore={presence}
      collaborationSession={collaborationSession}
    />,
  );

  await collaborationSession.connect(new AbortController().signal);

  expect(rendered.getByText('connected')).toHaveClass('tego-sheet__collaboration-status');
});
