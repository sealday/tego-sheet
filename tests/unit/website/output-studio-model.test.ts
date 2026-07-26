import { describe, expect, it } from 'vitest';
import type { Diagnostic, GeneratedDocument } from 'tego-sheet';
import {
  createOutputStudioState,
  groupOutputDiagnostics,
  hasBlockingDiagnostics,
  outputFilename,
  outputPipelineStages,
  reduceOutputStudioState,
} from '../../../website/src/components/playground/output-studio-model';

function diagnostic(severity: Diagnostic['severity']): Diagnostic {
  return {
    code: `TEST_${severity.toUpperCase()}`,
    severity,
    domain: 'template',
    stage: 'render',
    message: `${severity} diagnostic`,
  };
}

describe('Output Studio model', () => {
  it('marks output stale after a draft edit and disables generated output', () => {
    const ready = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      committedRevision: 1,
      generatedRevision: 1,
      generatedDocument: {} as GeneratedDocument,
      outputs: {
        ...createOutputStudioState().outputs,
        pdf: { requestId: 4, status: 'busy' as const, message: '' },
        xlsx: { requestId: 3, status: 'success' as const, message: 'XLSX downloaded' },
      },
    };

    expect(reduceOutputStudioState(ready, { type: 'draft-changed' })).toMatchObject({
      phase: 'dirty',
      committedRevision: 1,
      generatedRevision: 1,
      generatedDocument: ready.generatedDocument,
      outputs: expect.objectContaining({
        pdf: { requestId: null, status: 'cancelled', message: 'PDF generation cancelled.' },
        xlsx: { requestId: null, status: 'idle', message: '' },
      }),
    });
  });

  it('returns the same state for another draft edit while already dirty', () => {
    const dirty = {
      ...createOutputStudioState(),
      phase: 'dirty' as const,
      committedRevision: 1,
      generatedRevision: 1,
    };

    expect(reduceOutputStudioState(dirty, { type: 'draft-changed' })).toBe(dirty);
  });

  it('ignores render results from older revisions after a newer render starts', () => {
    const previousDocument = {} as GeneratedDocument;
    const rendering = reduceOutputStudioState(
      {
        ...createOutputStudioState(),
        phase: 'ready',
        committedRevision: 2,
        generatedRevision: 2,
        generatedDocument: previousDocument,
      },
      { type: 'render-started', revision: 3 },
    );

    expect(
      reduceOutputStudioState(rendering, {
        type: 'render-succeeded',
        revision: 2,
        document: {} as GeneratedDocument,
        metadata: {
          invoiceId: 'INV-2',
          title: 'Invoice',
          activePrintProfileId: 'invoice-a4',
          activePrintProfileName: 'Invoice · A4',
        },
        diagnostics: [],
      }),
    ).toBe(rendering);
    expect(
      reduceOutputStudioState(rendering, {
        type: 'render-blocked',
        revision: 2,
        diagnostics: [diagnostic('error')],
      }),
    ).toBe(rendering);
  });

  it('preserves the last generated document when the current render is blocked', () => {
    const previousDocument = {} as GeneratedDocument;
    const rendering = {
      ...createOutputStudioState(),
      phase: 'rendering' as const,
      committedRevision: 2,
      generatedRevision: 1,
      generatedDocument: previousDocument,
    };
    const error = diagnostic('error');

    expect(
      reduceOutputStudioState(rendering, {
        type: 'render-blocked',
        revision: 2,
        diagnostics: [error],
      }),
    ).toMatchObject({
      phase: 'blocked',
      committedRevision: 2,
      generatedRevision: 1,
      generatedDocument: previousDocument,
      diagnostics: [error],
    });
  });

  it('clears an older output operation before starting a new render', () => {
    const exporting = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      outputs: {
        ...createOutputStudioState().outputs,
        pdf: { requestId: 4, status: 'busy' as const, message: '' },
        xlsx: { requestId: 3, status: 'error' as const, message: 'Older output error' },
      },
    };

    expect(
      reduceOutputStudioState(exporting, { type: 'render-started', revision: 2 }),
    ).toMatchObject({
      phase: 'rendering',
      committedRevision: 2,
      outputs: expect.objectContaining({
        pdf: { requestId: null, status: 'cancelled', message: 'PDF generation cancelled.' },
        xlsx: { requestId: null, status: 'idle', message: '' },
      }),
    });
  });

  it('starts a reset revision with every output outcome idle', () => {
    const state = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      diagnostics: [diagnostic('error')],
      outputs: {
        print: { requestId: null, status: 'cancelled' as const, message: 'Print cancelled.' },
        pdf: { requestId: 4, status: 'busy' as const, message: 'Generating PDF…' },
        png: { requestId: 5, status: 'success' as const, message: 'PNG page 1 downloaded' },
        xlsx: { requestId: 6, status: 'error' as const, message: 'XLSX failed' },
      },
    };

    expect(reduceOutputStudioState(state, { type: 'reset-started', revision: 3 })).toMatchObject({
      phase: 'rendering',
      committedRevision: 3,
      diagnostics: [],
      outputs: {
        print: { requestId: null, status: 'idle', message: '' },
        pdf: { requestId: null, status: 'idle', message: '' },
        png: { requestId: null, status: 'idle', message: '' },
        xlsx: { requestId: null, status: 'idle', message: '' },
      },
    });
  });

  it('records output completion and failure without discarding generated output', () => {
    const document = {} as GeneratedDocument;
    const ready = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      committedRevision: 1,
      generatedRevision: 1,
      generatedDocument: document,
    };
    const exporting = reduceOutputStudioState(ready, {
      type: 'output-started',
      kind: 'xlsx',
      requestId: 7,
    });

    expect(exporting).toMatchObject({
      outputs: {
        xlsx: { requestId: 7, status: 'busy', message: 'Generating XLSX…' },
      },
      generatedDocument: document,
    });
    expect(
      reduceOutputStudioState(exporting, {
        type: 'output-failed',
        kind: 'xlsx',
        requestId: 7,
        message: 'XLSX output failed',
      }),
    ).toMatchObject({
      outputs: {
        xlsx: { requestId: 7, status: 'error', message: 'XLSX output failed' },
      },
      generatedDocument: document,
    });
    expect(
      reduceOutputStudioState(exporting, {
        type: 'output-finished',
        kind: 'xlsx',
        requestId: 7,
        message: 'XLSX downloaded',
      }),
    ).toMatchObject({
      outputs: {
        xlsx: { requestId: 7, status: 'success', message: 'XLSX downloaded' },
      },
      generatedDocument: document,
    });
  });

  it('retains outcomes per kind and ignores stale completions by request id', () => {
    const state = {
      ...createOutputStudioState(),
      outputs: {
        ...createOutputStudioState().outputs,
        pdf: { requestId: 9, status: 'busy' as const, message: '' },
        xlsx: { requestId: 3, status: 'error' as const, message: 'XLSX failed' },
      },
    };

    expect(
      reduceOutputStudioState(state, {
        type: 'output-finished',
        kind: 'pdf',
        requestId: 8,
        message: 'Stale PDF downloaded',
      }),
    ).toBe(state);
    expect(
      reduceOutputStudioState(state, {
        type: 'output-finished',
        kind: 'pdf',
        requestId: 9,
        message: 'PDF downloaded',
      }),
    ).toMatchObject({
      outputs: {
        pdf: { requestId: 9, status: 'success', message: 'PDF downloaded' },
        xlsx: { requestId: 3, status: 'error', message: 'XLSX failed' },
      },
    });
  });

  it('creates deterministic and sanitized output filenames', () => {
    expect(outputFilename('pdf', 'INV-2026-042')).toBe('invoice-INV-2026-042.pdf');
    expect(outputFilename('png', 'INV-2026-042', 0)).toBe('invoice-INV-2026-042-page-1.png');
    expect(outputFilename('xlsx', 'INV-2026-042')).toBe('invoice-INV-2026-042.xlsx');
    expect(outputFilename('pdf', ' Invoice / 42 ')).toBe('invoice--Invoice-42-.pdf');
  });

  it('treats only error diagnostics as blocking', () => {
    expect(hasBlockingDiagnostics([diagnostic('info'), diagnostic('warning')])).toBe(false);
    expect(hasBlockingDiagnostics([diagnostic('warning'), diagnostic('error')])).toBe(true);
  });

  it('groups blocking diagnostics separately from warnings', () => {
    const warning = diagnostic('warning');
    const error = diagnostic('error');

    expect(groupOutputDiagnostics([warning, diagnostic('info'), error])).toEqual({
      blocking: [error],
      warnings: [warning],
    });
  });

  it('exposes compile, bind, and paginate state for ready and blocked revisions', () => {
    const ready = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      committedRevision: 2,
      generatedRevision: 2,
    };
    expect(outputPipelineStages(ready)).toEqual([
      { id: 'compile', label: 'Compile', status: 'complete' },
      { id: 'bind', label: 'Bind', status: 'complete' },
      { id: 'paginate', label: 'Paginate', status: 'complete' },
    ]);

    expect(
      outputPipelineStages({
        ...ready,
        phase: 'blocked',
        generatedRevision: 1,
        diagnostics: [
          {
            ...diagnostic('error'),
            stage: 'expand',
          },
        ],
      }),
    ).toEqual([
      { id: 'compile', label: 'Compile', status: 'complete' },
      { id: 'bind', label: 'Bind', status: 'blocked' },
      { id: 'paginate', label: 'Paginate', status: 'pending' },
    ]);
  });
});
