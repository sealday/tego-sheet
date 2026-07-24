import type { SpreadsheetDocument } from '../document';
import { XlsxAdapter } from '../output/xlsx';
import type { GeneratedDocument, TemplatePrintProfile } from '../template';
import {
  exportResult,
  InterchangeError,
  resolveLimits,
  throwIfAborted,
  type InterchangeLimits,
  type InterchangeWriteOptions,
  type WorkbookExportResult,
  type WorkbookWriter,
} from './contracts';

function defaultPrintProfile(document: SpreadsheetDocument): TemplatePrintProfile {
  return {
    id: 'interchange-default-print',
    name: 'Workbook',
    targets: document.workbook.sheets.map((sheet) => ({ type: 'sheet', sheetId: sheet.id })),
    page: {
      paper: { type: 'A4' },
      orientation: 'portrait',
      margins: { top: 48, right: 48, bottom: 48, left: 48 },
      scale: { type: 'fixed', value: 1 },
    },
    manualBreaks: [],
    showGridlines: true,
    showHeadings: true,
  };
}

function generatedDocument(document: SpreadsheetDocument): GeneratedDocument {
  const profile =
    document.templates.flatMap((template) => template.printProfiles)[0] ??
    defaultPrintProfile(document);
  return {
    workbook: document.workbook,
    calculatedCells: [],
    worksheets: document.workbook.sheets.map((sheet) => ({
      sheetId: sheet.id,
      visibility: sheet.visibility,
      conditionalFormatting: sheet.conditionalFormatting,
    })),
    print: { pages: [], displayList: { pages: [], diagnostics: [] }, profile },
    resources: {
      byHash: {},
      byReference: {},
      totalBytes: 0,
      dispose: async () => undefined,
    },
    objects: [],
    diagnostics: [],
    metadata: {
      templateId: document.templates[0]?.id ?? 'interchange-default-template',
      profileId: profile.id,
      sourceDocumentHash: 'interchange-direct-write',
      locale: document.workbook.settings.localeHint ?? 'en-US',
      timeZone: 'UTC',
      generatedAt: '1970-01-01T00:00:00.000Z',
    },
  };
}

/** Creates a deterministic bounded XLSX writer backed by the canonical OOXML adapter. */
export function createXlsxWriter(configuredLimits: InterchangeLimits = {}): WorkbookWriter {
  const limits = resolveLimits(configuredLimits);
  const adapter = new XlsxAdapter({
    limits: {
      maxCells: limits.maxCells,
      maxPackageBytes: Math.min(limits.maxPackageBytes, limits.maxOutputBytes),
      maxUncompressedBytes: limits.maxUncompressedBytes,
    },
  });
  const writeResult = async (
    document: SpreadsheetDocument,
    options: InterchangeWriteOptions = {},
  ): Promise<WorkbookExportResult> => {
    throwIfAborted(options.signal);
    const unsupported = document.workbook.sheets.flatMap((sheet) =>
      sheet.objects.length === 0 ? [] : ['xlsx:drawing-objects'],
    );
    const blob = await adapter.render(generatedDocument(document), {
      formulaMode: 'formula-and-cached-value',
      compatibility: 'excel',
      signal: options.signal,
    });
    throwIfAborted(options.signal);
    const maxOutputBytes = options.maxOutputBytes ?? limits.maxOutputBytes;
    if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 0) {
      throw new TypeError('maxOutputBytes must be a non-negative safe integer');
    }
    if (blob.size > maxOutputBytes || blob.size > limits.maxPackageBytes) {
      throw new InterchangeError('OUTPUT_LIMIT_EXCEEDED', 'XLSX output byte limit exceeded');
    }
    return exportResult('xlsx', blob, unsupported);
  };
  return Object.freeze({
    format: 'xlsx',
    async write(document: SpreadsheetDocument, options = {}) {
      return (await writeResult(document, options as InterchangeWriteOptions)).blob;
    },
    writeResult(document: SpreadsheetDocument, options = {}) {
      return writeResult(document, options as InterchangeWriteOptions);
    },
  });
}
