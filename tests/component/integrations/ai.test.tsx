import { fireEvent, render } from '@testing-library/react';
import { expect, it, vi } from 'vitest';
import { AiProposalPanel } from '../../../src';

it('requires an explicit proposal decision and settles only the selected action', () => {
  const accept = vi.fn(() => ({ status: 'committed' }));
  const reject = vi.fn();
  const onAccepted = vi.fn();
  const rendered = render(
    <AiProposalPanel
      session={{
        proposal: {
          id: 'proposal-1',
          summary: 'Normalize the selected values',
          assumptions: ['Headers are in row 1'],
          commands: [],
        },
        preview: { status: 'ready', diagnostics: [] },
        contextSummary: {
          sheetCount: 1,
          cellCount: 4,
          omittedCellCount: 10,
          serializedBytes: 256,
        },
        accept,
        reject,
      }}
      onAccepted={onAccepted}
    />,
  );

  expect(accept).not.toHaveBeenCalled();
  expect(rendered.getByText('1 sheets, 4 cells, 256 bytes')).not.toBeNull();
  fireEvent.click(rendered.getByRole('button', { name: 'Apply' }));

  expect(accept).toHaveBeenCalledTimes(1);
  expect(reject).not.toHaveBeenCalled();
  expect(onAccepted).toHaveBeenCalledWith({ status: 'committed' });
  expect((rendered.getByRole('button', { name: 'Apply' }) as HTMLButtonElement).disabled).toBe(
    true,
  );
  expect((rendered.getByRole('button', { name: 'Reject' }) as HTMLButtonElement).disabled).toBe(
    true,
  );
});
