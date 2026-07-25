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

  const rendered = await renderSpreadsheetTemplate(
    {
      template: compilation.template,
      currentDocumentHash: hashSpreadsheetDocument(sourceDocument),
      data: request.data,
      profileId: request.template.printProfiles[0]!.id,
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
