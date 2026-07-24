import { useState } from 'react';

/** Props for the explicit AI proposal review surface. */
export interface AiProposalPanelProps<ApplyResult = unknown> {
  /** Validated proposal, dry-run preview, and the one-shot accept/reject boundary. */
  readonly session: {
    readonly proposal: {
      readonly id: string;
      readonly summary: string;
      readonly assumptions: readonly string[];
      readonly commands: readonly unknown[];
    };
    readonly contextSummary?: {
      readonly sheetCount: number;
      readonly cellCount: number;
      readonly omittedCellCount: number;
      readonly serializedBytes: number;
    };
    readonly preview: {
      readonly status: 'ready' | 'noop';
      readonly diagnostics: readonly { readonly severity?: string }[];
      readonly baseRevision?: number;
      readonly document?: unknown;
    };
    accept(): ApplyResult;
    reject(): void;
  };
  /** Receives the result returned by the accepted proposal session. */
  readonly onAccepted?: (result: ApplyResult) => void;
  /** Runs after the proposal session is rejected. */
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
