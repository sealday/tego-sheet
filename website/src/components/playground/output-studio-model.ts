import type { Diagnostic, GeneratedDocument } from 'tego-sheet';

export type OutputStudioPhase = 'ready' | 'dirty' | 'rendering' | 'blocked';
export type OutputKind = 'print' | 'pdf' | 'png' | 'xlsx';
export type OutputStatus = 'idle' | 'busy' | 'success' | 'error';

export interface OutputState {
  readonly requestId: number | null;
  readonly status: OutputStatus;
  readonly message: string;
}

export interface OutputDocumentMetadata {
  readonly invoiceId: string;
  readonly title: string;
}

export interface OutputStudioState {
  readonly phase: OutputStudioPhase;
  readonly committedRevision: number;
  readonly generatedRevision: number | null;
  readonly generatedDocument: GeneratedDocument | null;
  readonly generatedMetadata: OutputDocumentMetadata | null;
  readonly diagnostics: readonly Diagnostic[];
  readonly outputs: Readonly<Record<OutputKind, OutputState>>;
}

export type OutputStudioAction =
  | { readonly type: 'draft-changed' }
  | { readonly type: 'outputs-cancelled' }
  | { readonly type: 'render-started'; readonly revision: number }
  | {
      readonly type: 'render-succeeded';
      readonly revision: number;
      readonly document: GeneratedDocument;
      readonly metadata: OutputDocumentMetadata;
      readonly diagnostics: readonly Diagnostic[];
    }
  | {
      readonly type: 'render-blocked';
      readonly revision: number;
      readonly diagnostics: readonly Diagnostic[];
    }
  | { readonly type: 'output-started'; readonly kind: OutputKind; readonly requestId: number }
  | {
      readonly type: 'output-finished';
      readonly kind: OutputKind;
      readonly requestId: number;
      readonly message: string;
    }
  | {
      readonly type: 'output-failed';
      readonly kind: OutputKind;
      readonly requestId: number;
      readonly message: string;
    };

function idleOutput(): OutputState {
  return { requestId: null, status: 'idle', message: '' };
}

function createOutputs(): Readonly<Record<OutputKind, OutputState>> {
  return {
    print: idleOutput(),
    pdf: idleOutput(),
    png: idleOutput(),
    xlsx: idleOutput(),
  };
}

function cancelBusyOutputs(
  outputs: Readonly<Record<OutputKind, OutputState>>,
): Readonly<Record<OutputKind, OutputState>> {
  const cancelled = (output: OutputState): OutputState =>
    output.status === 'busy' ? idleOutput() : output;
  return {
    print: cancelled(outputs.print),
    pdf: cancelled(outputs.pdf),
    png: cancelled(outputs.png),
    xlsx: cancelled(outputs.xlsx),
  };
}

export function createOutputStudioState(): OutputStudioState {
  return {
    phase: 'rendering',
    committedRevision: 0,
    generatedRevision: null,
    generatedDocument: null,
    generatedMetadata: null,
    diagnostics: [],
    outputs: createOutputs(),
  };
}

export function reduceOutputStudioState(
  state: OutputStudioState,
  action: OutputStudioAction,
): OutputStudioState {
  if (action.type === 'render-started') {
    return {
      ...state,
      phase: 'rendering',
      committedRevision: action.revision,
      outputs: cancelBusyOutputs(state.outputs),
    };
  }
  if ('revision' in action && action.revision !== state.committedRevision) return state;
  if (action.type === 'draft-changed' && state.phase === 'dirty') return state;

  switch (action.type) {
    case 'draft-changed':
      return { ...state, phase: 'dirty', outputs: cancelBusyOutputs(state.outputs) };
    case 'outputs-cancelled':
      return { ...state, outputs: cancelBusyOutputs(state.outputs) };
    case 'render-succeeded':
      return {
        ...state,
        phase: 'ready',
        generatedRevision: action.revision,
        generatedDocument: action.document,
        generatedMetadata: action.metadata,
        diagnostics: action.diagnostics,
      };
    case 'render-blocked':
      return { ...state, phase: 'blocked', diagnostics: action.diagnostics };
    case 'output-started':
      return {
        ...state,
        outputs: {
          ...state.outputs,
          [action.kind]: { requestId: action.requestId, status: 'busy', message: '' },
        },
      };
    case 'output-finished':
    case 'output-failed': {
      if (state.outputs[action.kind].requestId !== action.requestId) return state;
      return {
        ...state,
        outputs: {
          ...state.outputs,
          [action.kind]: {
            requestId: action.requestId,
            status: action.type === 'output-finished' ? 'success' : 'error',
            message: action.message,
          },
        },
      };
    }
  }
}

export function outputFilename(
  kind: Exclude<OutputKind, 'print'>,
  invoiceId: string,
  page = 0,
): string {
  const safeId = invoiceId.replace(/[^A-Za-z0-9._-]+/g, '-');
  return kind === 'png' ? `invoice-${safeId}-page-${page + 1}.png` : `invoice-${safeId}.${kind}`;
}

export function hasBlockingDiagnostics(diagnostics: readonly Diagnostic[]): boolean {
  return diagnostics.some(({ severity }) => severity === 'error');
}
