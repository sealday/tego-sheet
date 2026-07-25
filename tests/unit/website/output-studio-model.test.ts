import { describe, expect, it } from 'vitest';
import type { Diagnostic, GeneratedDocument } from 'tego-sheet';
import {
  createOutputStudioState,
  hasBlockingDiagnostics,
  outputFilename,
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
      activeOutput: 'pdf' as const,
    };

    expect(reduceOutputStudioState(ready, { type: 'draft-changed' })).toMatchObject({
      phase: 'dirty',
      committedRevision: 1,
      generatedRevision: 1,
      generatedDocument: ready.generatedDocument,
      activeOutput: null,
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
      activeOutput: 'pdf' as const,
      outputMessage: 'Older output error',
    };

    expect(
      reduceOutputStudioState(exporting, { type: 'render-started', revision: 2 }),
    ).toMatchObject({
      phase: 'rendering',
      committedRevision: 2,
      activeOutput: null,
      outputMessage: '',
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
    });

    expect(exporting).toMatchObject({
      activeOutput: 'xlsx',
      outputMessage: '',
      generatedDocument: document,
    });
    expect(
      reduceOutputStudioState(exporting, {
        type: 'output-failed',
        message: 'XLSX output failed',
      }),
    ).toMatchObject({
      activeOutput: null,
      outputMessage: 'XLSX output failed',
      generatedDocument: document,
    });
    expect(
      reduceOutputStudioState(exporting, {
        type: 'output-finished',
        message: 'XLSX downloaded',
      }),
    ).toMatchObject({
      activeOutput: null,
      outputMessage: 'XLSX downloaded',
      generatedDocument: document,
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
});
