import type { Diagnostic, GeneratedDocument } from 'tego-sheet';

export type OutputStudioPhase = 'ready' | 'dirty' | 'rendering' | 'blocked';
export type OutputKind = 'print' | 'pdf' | 'png' | 'xlsx';
export type OutputStatus = 'idle' | 'busy' | 'success' | 'error' | 'cancelled';
export type OutputPipelineStageId = 'compile' | 'bind' | 'paginate';
export type OutputPipelineStageStatus = 'pending' | 'active' | 'complete' | 'blocked';

export interface OutputPipelineStage {
  readonly id: OutputPipelineStageId;
  readonly label: 'Compile' | 'Bind' | 'Paginate';
  readonly status: OutputPipelineStageStatus;
}

export interface GroupedOutputDiagnostics {
  readonly blocking: readonly Diagnostic[];
  readonly warnings: readonly Diagnostic[];
}

export interface OutputState {
  readonly requestId: number | null;
  readonly status: OutputStatus;
  readonly message: string;
}

export interface OutputDocumentMetadata {
  readonly invoiceId: string;
  readonly title: string;
  readonly activePrintProfileId: string;
  readonly activePrintProfileName: string;
}

export interface OutputStudioState {
  readonly phase: OutputStudioPhase;
  readonly committedRevision: number;
  readonly generatedRevision: number | null;
  readonly generatedDocument: GeneratedDocument | null;
  readonly generatedMetadata: OutputDocumentMetadata | null;
  readonly diagnostics: readonly Diagnostic[];
  readonly pipelineStages: readonly OutputPipelineStage[];
  readonly outputs: Readonly<Record<OutputKind, OutputState>>;
}

export type OutputStudioAction =
  | { readonly type: 'draft-changed' }
  | { readonly type: 'outputs-cancelled' }
  | { readonly type: 'render-started'; readonly revision: number }
  | { readonly type: 'reset-started'; readonly revision: number }
  | {
      readonly type: 'render-progress';
      readonly revision: number;
      readonly stage: OutputPipelineStageId;
      readonly status: 'active' | 'blocked';
    }
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

const PIPELINE_STAGE_DEFINITIONS = Object.freeze([
  { id: 'compile', label: 'Compile' },
  { id: 'bind', label: 'Bind' },
  { id: 'paginate', label: 'Paginate' },
] as const);

function pipelineStages(
  currentStage?: OutputPipelineStageId,
  currentStatus: 'active' | 'blocked' = 'active',
): readonly OutputPipelineStage[] {
  const currentIndex =
    currentStage === undefined
      ? -1
      : PIPELINE_STAGE_DEFINITIONS.findIndex(({ id }) => id === currentStage);
  return PIPELINE_STAGE_DEFINITIONS.map(({ id, label }, index) => ({
    id,
    label,
    status:
      currentIndex < 0
        ? 'pending'
        : index < currentIndex
          ? 'complete'
          : index === currentIndex
            ? currentStatus
            : 'pending',
  }));
}

function completedPipelineStages(): readonly OutputPipelineStage[] {
  return PIPELINE_STAGE_DEFINITIONS.map(({ id, label }) => ({
    id,
    label,
    status: 'complete',
  }));
}

function outputLabel(kind: OutputKind): string {
  if (kind === 'print') return 'Print';
  return kind === 'png' ? 'PNG' : kind.toUpperCase();
}

function busyOutput(kind: OutputKind, requestId: number): OutputState {
  return {
    requestId,
    status: 'busy',
    message: kind === 'print' ? 'Opening print dialog…' : `Generating ${outputLabel(kind)}…`,
  };
}

function cancelledOutput(kind: OutputKind): OutputState {
  return {
    requestId: null,
    status: 'cancelled',
    message: kind === 'print' ? 'Print cancelled.' : `${outputLabel(kind)} generation cancelled.`,
  };
}

function cancelBusyOutputs(
  outputs: Readonly<Record<OutputKind, OutputState>>,
): Readonly<Record<OutputKind, OutputState>> {
  const cancelled = (kind: OutputKind): OutputState =>
    outputs[kind].status === 'busy' ? cancelledOutput(kind) : outputs[kind];
  return {
    print: cancelled('print'),
    pdf: cancelled('pdf'),
    png: cancelled('png'),
    xlsx: cancelled('xlsx'),
  };
}

function changeRevisionOutputs(
  outputs: Readonly<Record<OutputKind, OutputState>>,
): Readonly<Record<OutputKind, OutputState>> {
  const changed = (kind: OutputKind): OutputState =>
    outputs[kind].status === 'busy' ? cancelledOutput(kind) : idleOutput();
  return {
    print: changed('print'),
    pdf: changed('pdf'),
    png: changed('png'),
    xlsx: changed('xlsx'),
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
    pipelineStages: pipelineStages(),
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
      diagnostics: [],
      pipelineStages: pipelineStages(),
      outputs: changeRevisionOutputs(state.outputs),
    };
  }
  if (action.type === 'reset-started') {
    return {
      ...state,
      phase: 'rendering',
      committedRevision: action.revision,
      diagnostics: [],
      pipelineStages: pipelineStages(),
      outputs: createOutputs(),
    };
  }
  if ('revision' in action && action.revision !== state.committedRevision) return state;
  if (action.type === 'draft-changed' && state.phase === 'dirty') return state;

  switch (action.type) {
    case 'draft-changed':
      return {
        ...state,
        phase: 'dirty',
        pipelineStages: pipelineStages(),
        outputs: changeRevisionOutputs(state.outputs),
      };
    case 'render-progress':
      if (state.phase !== 'rendering') return state;
      return {
        ...state,
        pipelineStages: pipelineStages(action.stage, action.status),
      };
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
        pipelineStages: completedPipelineStages(),
      };
    case 'render-blocked': {
      const alreadyBlocked = state.pipelineStages.some(({ status }) => status === 'blocked');
      const blockingDiagnostic = action.diagnostics.find(({ severity }) => severity === 'error');
      return {
        ...state,
        phase: 'blocked',
        diagnostics: action.diagnostics,
        pipelineStages:
          alreadyBlocked || blockingDiagnostic === undefined
            ? state.pipelineStages
            : pipelineStages(outputDiagnosticStage(blockingDiagnostic), 'blocked'),
      };
    }
    case 'output-started':
      return {
        ...state,
        outputs: {
          ...state.outputs,
          [action.kind]: busyOutput(action.kind, action.requestId),
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

export function groupOutputDiagnostics(
  diagnostics: readonly Diagnostic[],
): GroupedOutputDiagnostics {
  return {
    blocking: diagnostics.filter(({ severity }) => severity === 'error'),
    warnings: diagnostics.filter(({ severity }) => severity === 'warning'),
  };
}

export function outputDiagnosticStage(diagnostic: Diagnostic): OutputPipelineStageId {
  if (
    diagnostic.stage === 'decode' ||
    diagnostic.stage === 'validate' ||
    diagnostic.stage === 'plan' ||
    diagnostic.stage === 'compile'
  ) {
    return 'compile';
  }
  if (
    diagnostic.stage === 'resolve' ||
    diagnostic.stage === 'expand' ||
    diagnostic.stage === 'recalculate'
  ) {
    return 'bind';
  }
  return 'paginate';
}

export function outputPipelineStages(state: OutputStudioState): readonly OutputPipelineStage[] {
  return state.pipelineStages;
}
