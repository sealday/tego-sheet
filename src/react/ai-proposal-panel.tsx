import { useState } from 'react';

/** Props for the explicit AI proposal review surface. */
export interface AiProposalPanelProps<ApplyResult = unknown> {
  /** Validated proposal, dry-run preview, and the one-shot accept/reject boundary. */
  readonly session: {
    /** Validated command proposal awaiting an explicit user decision. */
    readonly proposal: {
      /** Stable proposal identifier. */
      readonly id: string;
      /** Human-readable description of the proposed change. */
      readonly summary: string;
      /** Model assumptions displayed before the user decides. */
      readonly assumptions: readonly string[];
      /** Validated opaque commands retained by the proposal session. */
      readonly commands: readonly unknown[];
    };
    /** Value-free scope summary for the sanitized AI context. */
    readonly contextSummary?: {
      /** Number of worksheets represented in the context. */
      readonly sheetCount: number;
      /** Number of cells represented in the context. */
      readonly cellCount: number;
      /** Number of cells omitted by limits or sanitization. */
      readonly omittedCellCount: number;
      /** Serialized byte size of the sanitized context. */
      readonly serializedBytes: number;
    };
    /** Dry-run result shown before any command can be applied. */
    readonly preview: {
      /** Whether the proposal produces a change or is a no-op. */
      readonly status: 'ready' | 'noop';
      /** Diagnostics produced by validation and the dry run. */
      readonly diagnostics: readonly {
        /** Optional host-defined diagnostic severity. */
        readonly severity?: string;
      }[];
      /** Document revision against which the proposal was previewed. */
      readonly baseRevision?: number;
      /** Optional isolated preview document. */
      readonly document?: unknown;
    };
    /** Applies the already validated proposal exactly once. */
    accept(): ApplyResult;
    /** Rejects the proposal without applying commands. */
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
