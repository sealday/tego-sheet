import {
  compileSpreadsheetTemplate,
  hashSpreadsheetDocument,
  renderSpreadsheetTemplate,
  type Diagnostic,
  type GeneratedDocument,
  type RenderEnvironment,
  type SpreadsheetDocument,
  type SpreadsheetTemplate,
} from 'tego-sheet';

export interface OutputRevisionRequest {
  readonly revision: number;
  readonly document: SpreadsheetDocument;
  readonly template: SpreadsheetTemplate;
  readonly activePrintProfileId: string;
  readonly data: unknown;
  readonly environment: RenderEnvironment;
  readonly signal: AbortSignal;
  readonly onProgress?: (progress: OutputRevisionProgress) => void;
}

export interface OutputRevisionProgress {
  readonly revision: number;
  readonly stage: 'compile' | 'bind' | 'paginate';
  readonly status: 'active' | 'blocked';
}

export interface OutputRevisionResult {
  readonly revision: number;
  readonly diagnostics: readonly Diagnostic[];
  readonly document?: GeneratedDocument;
}

export async function renderOutputRevision(
  request: OutputRevisionRequest,
): Promise<OutputRevisionResult> {
  let activeStage: OutputRevisionProgress['stage'] = 'compile';
  const report = (
    stage: OutputRevisionProgress['stage'],
    status: OutputRevisionProgress['status'],
  ): void => {
    if (request.signal.aborted) return;
    activeStage = stage;
    request.onProgress?.(Object.freeze({ revision: request.revision, stage, status }));
  };
  report('compile', 'active');
  const sourceDocument: SpreadsheetDocument = {
    ...request.document,
    templates: [
      ...request.document.templates.filter(({ id }) => id !== request.template.id),
      request.template,
    ],
  };
  const compilation = compileSpreadsheetTemplate(sourceDocument, request.template);
  if (compilation.template === undefined) {
    report('compile', 'blocked');
    return Object.freeze({
      revision: request.revision,
      diagnostics: compilation.diagnostics,
    });
  }

  const profile = request.template.printProfiles.find(
    ({ id }) => id === request.activePrintProfileId,
  );
  if (profile === undefined) {
    const diagnostic: Diagnostic = Object.freeze({
      code: 'INVALID_PRINT_TARGET',
      severity: 'error',
      domain: 'template',
      stage: 'validate',
      message:
        request.template.printProfiles.length === 0
          ? 'Template has no print profile'
          : 'Selected print profile no longer exists',
    });
    report('compile', 'blocked');
    return Object.freeze({
      revision: request.revision,
      diagnostics: Object.freeze([...compilation.diagnostics, diagnostic]),
    });
  }

  const renderRequest = {
    template: compilation.template,
    currentDocumentHash: hashSpreadsheetDocument(sourceDocument),
    data: request.data,
    profileId: profile.id,
    missingValue: 'error' as const,
    signal: request.signal,
    __internalStageProgress(stage: 'bind' | 'paginate') {
      report(stage, 'active');
    },
  };
  const rendered = await renderSpreadsheetTemplate(renderRequest, request.environment);
  if (
    rendered.document === undefined &&
    rendered.diagnostics.some(({ severity }) => severity === 'error')
  ) {
    report(activeStage, 'blocked');
  }

  return Object.freeze({
    revision: request.revision,
    diagnostics: Object.freeze([...compilation.diagnostics, ...rendered.diagnostics]),
    ...(rendered.document === undefined ? {} : { document: rendered.document }),
  });
}
