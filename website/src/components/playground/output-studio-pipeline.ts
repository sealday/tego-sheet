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
}

export interface OutputRevisionResult {
  readonly revision: number;
  readonly diagnostics: readonly Diagnostic[];
  readonly document?: GeneratedDocument;
}

export async function renderOutputRevision(
  request: OutputRevisionRequest,
): Promise<OutputRevisionResult> {
  const sourceDocument: SpreadsheetDocument = {
    ...request.document,
    templates: [
      ...request.document.templates.filter(({ id }) => id !== request.template.id),
      request.template,
    ],
  };
  const compilation = compileSpreadsheetTemplate(sourceDocument, request.template);
  if (compilation.template === undefined) {
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
    return Object.freeze({
      revision: request.revision,
      diagnostics: Object.freeze([...compilation.diagnostics, diagnostic]),
    });
  }

  const rendered = await renderSpreadsheetTemplate(
    {
      template: compilation.template,
      currentDocumentHash: hashSpreadsheetDocument(sourceDocument),
      data: request.data,
      profileId: profile.id,
      missingValue: 'error',
      signal: request.signal,
    },
    request.environment,
  );

  return Object.freeze({
    revision: request.revision,
    diagnostics: Object.freeze([...compilation.diagnostics, ...rendered.diagnostics]),
    ...(rendered.document === undefined ? {} : { document: rendered.document }),
  });
}
