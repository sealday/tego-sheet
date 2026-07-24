import { useState } from 'react';
import type { AiContextSummary, AiProposalSession, AiTransactionPreview } from '../integrations/ai';

export interface AiProposalPanelProps<ApplyResult = unknown> {
  readonly session: AiProposalSession<ApplyResult> & {
    readonly contextSummary?: AiContextSummary;
    readonly preview: AiTransactionPreview;
  };
  readonly onAccepted?: (result: ApplyResult) => void;
  readonly onRejected?: () => void;
}

/** Explicit review surface for a validated, dry-run AI command proposal. */
export function AiProposalPanel<ApplyResult>({
  session,
  onAccepted,
  onRejected,
}: AiProposalPanelProps<ApplyResult>) {
  const [settled, setSettled] = useState(false);
  const [error, setError] = useState<string>();
  const accept = (): void => {
    try {
      const result = session.accept();
      setSettled(true);
      onAccepted?.(result);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'AI proposal could not be applied');
    }
  };
  const reject = (): void => {
    session.reject();
    setSettled(true);
    onRejected?.();
  };

  return (
    <section className="tego-sheet__ai-proposal" aria-label="AI proposal review">
      <h2>{session.proposal.summary}</h2>
      {session.contextSummary === undefined ? null : (
        <p className="tego-sheet__ai-context-summary">
          {session.contextSummary.sheetCount} sheets, {session.contextSummary.cellCount} cells,{' '}
          {session.contextSummary.serializedBytes} bytes
        </p>
      )}
      <p>Preview: {session.preview.status}</p>
      {session.proposal.assumptions.length === 0 ? null : (
        <ul>
          {session.proposal.assumptions.map((assumption) => (
            <li key={assumption}>{assumption}</li>
          ))}
        </ul>
      )}
      {error === undefined ? null : <p role="alert">{error}</p>}
      <button type="button" disabled={settled} onClick={accept}>
        Apply
      </button>
      <button type="button" disabled={settled} onClick={reject}>
        Reject
      </button>
    </section>
  );
}
