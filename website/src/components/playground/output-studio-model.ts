import type { Diagnostic, GeneratedDocument } from 'tego-sheet';

export type OutputStudioPhase = 'ready' | 'dirty' | 'rendering' | 'blocked';
export type OutputKind = 'print' | 'pdf' | 'png' | 'xlsx';

export interface OutputStudioState {
  readonly phase: OutputStudioPhase;
  readonly committedRevision: number;
  readonly generatedRevision: number | null;
  readonly generatedDocument: GeneratedDocument | null;
  readonly diagnostics: readonly Diagnostic[];
  readonly activeOutput: OutputKind | null;
  readonly outputMessage: string;
}

export type OutputStudioAction =
  | { readonly type: 'draft-changed' }
  | { readonly type: 'render-started'; readonly revision: number }
  | {
      readonly type: 'render-succeeded';
      readonly revision: number;
      readonly document: GeneratedDocument;
      readonly diagnostics: readonly Diagnostic[];
    }
  | {
      readonly type: 'render-blocked';
      readonly revision: number;
      readonly diagnostics: readonly Diagnostic[];
    }
  | { readonly type: 'output-started'; readonly kind: OutputKind }
  | { readonly type: 'output-finished'; readonly message: string }
  | { readonly type: 'output-failed'; readonly message: string };

export function createOutputStudioState(): OutputStudioState {
  return {
    phase: 'rendering',
    committedRevision: 0,
    generatedRevision: null,
    generatedDocument: null,
    diagnostics: [],
    activeOutput: null,
    outputMessage: '',
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
      activeOutput: null,
      outputMessage: '',
    };
  }
  if ('revision' in action && action.revision !== state.committedRevision) return state;
  if (action.type === 'draft-changed' && state.phase === 'dirty') return state;

  switch (action.type) {
    case 'draft-changed':
      return { ...state, phase: 'dirty', activeOutput: null };
    case 'render-succeeded':
      return {
        ...state,
        phase: 'ready',
        generatedRevision: action.revision,
        generatedDocument: action.document,
        diagnostics: action.diagnostics,
      };
    case 'render-blocked':
      return { ...state, phase: 'blocked', diagnostics: action.diagnostics };
    case 'output-started':
      return { ...state, activeOutput: action.kind, outputMessage: '' };
    case 'output-finished':
    case 'output-failed':
      return { ...state, activeOutput: null, outputMessage: action.message };
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
