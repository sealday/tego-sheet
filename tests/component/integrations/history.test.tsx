import { cleanup, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TegoSheet, type TegoSheetHandle } from '../../../src';
import { loadHistoryPreview } from '../../../src/integrations/history';
import { createPermissionSnapshot } from '../../../src/integrations/permission';
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

it('mounts a loaded version preview through the ordinary isolated readonly surface', async () => {
  const document = testDocument([{ name: 'Historical' }]);
  const preview = await loadHistoryPreview(
    {
      load: async () => ({
        id: 'version-1',
        documentId: document.id,
        revision: 'revision-1',
        document,
      }),
    },
    document.id,
    'version-1',
    createPermissionSnapshot({
      revision: 'permission-1',
      actorId: 'actor-1',
      grants: [
        {
          action: 'history:view',
          target: { type: 'document', documentId: document.id },
        },
      ],
    }),
    new AbortController().signal,
  );
  const ref = createRef<TegoSheetHandle>();
  render(<TegoSheet ref={ref} document={preview.document} readOnly={preview.readOnly} />);
  await waitFor(() => expect(ref.current).not.toBeNull());
  const sheet = document.workbook.sheets[0]!.id as never;

  expect(() => ref.current!.setCellText({ sheet, row: 0, column: 0 }, 'blocked')).toThrow(
    /read-only/u,
  );
});
