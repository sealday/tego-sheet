import type { SheetObject, SpreadsheetDocument } from '../document';
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

function dataUrlBytes(
  url: string | undefined,
  mimeType: string | undefined,
): readonly number[] | undefined {
  if (url === undefined || mimeType === undefined) return undefined;
  const match = /^data:([^;,]+);base64,([\s\S]*)$/i.exec(url);
  if (match === null || match[1]?.toLowerCase() !== mimeType.toLowerCase()) return undefined;
  try {
    return Object.freeze(
      atob(match[2]!)
        .split('')
        .map((character) => character.charCodeAt(0)),
    );
  } catch {
    return undefined;
  }
}

function objectRange(
  sheetId: SpreadsheetDocument['workbook']['sheets'][number]['id'],
  object: SheetObject,
) {
  if (object.anchor.type === 'two-cell') {
    return {
      sheetId,
      start: { row: object.anchor.from.row, column: object.anchor.from.column },
      end: { row: object.anchor.to.row, column: object.anchor.to.column },
    };
  }
  if (object.anchor.type === 'one-cell') {
    return {
      sheetId,
      start: { row: object.anchor.cell.row, column: object.anchor.cell.column },
      end: { row: object.anchor.cell.row, column: object.anchor.cell.column },
    };
  }
  return {
    sheetId,
    start: { row: 0, column: 0 },
    end: { row: 0, column: 0 },
  };
}

function generatedDocument(document: SpreadsheetDocument): GeneratedDocument {
  const profile =
    document.templates.flatMap((template) => template.printProfiles)[0] ??
    defaultPrintProfile(document);
  const resolved = document.resources.items.flatMap((resource) => {
    const bytes = dataUrlBytes(resource.url, resource.mimeType);
    return bytes === undefined || resource.mimeType === undefined
      ? []
      : [
          {
            reference: resource.id,
            resource: {
              contentHash: `xlsx:${resource.id}`,
              type: 'image' as const,
              mimeType: resource.mimeType,
              bytes,
            },
          },
        ];
  });
  const byReference = Object.fromEntries(
    resolved.map(({ reference, resource }) => [reference, resource]),
  );
  const byHash = Object.fromEntries(
    resolved.map(({ resource }) => [resource.contentHash, resource]),
  );
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
      byHash,
      byReference,
      totalBytes: resolved.reduce((total, { resource }) => total + resource.bytes.length, 0),
      dispose: async () => undefined,
    },
    objects: document.workbook.sheets.flatMap((sheet) =>
      sheet.objects.map((object) => {
        const range = objectRange(sheet.id, object);
        return {
          objectId: object.id,
          ...(object.kind === 'image' ? { resourceId: object.resourceId } : {}),
          policy: 'shared' as const,
          itemIndex: 0,
          source: range,
          generated: range,
        };
      }),
    ),
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
    const unsupported: string[] = [];
    if (document.workbook.sheets.some((sheet) => sheet.charts.length > 0)) {
      unsupported.push('xlsx:charts');
    }
    if (document.workbook.sheets.some((sheet) => sheet.sparklines.length > 0)) {
      unsupported.push('xlsx:sparklines');
    }
    if (
      document.workbook.sheets.some((sheet) =>
        sheet.tables.some((table) => table.columns.some((column) => column.dataType !== undefined)),
      )
    ) {
      unsupported.push('xlsx:table-column-data-type');
    }
    if (
      document.workbook.sheets.some((sheet) =>
        sheet.tables.some((table) => table.autoExpand !== undefined),
      )
    ) {
      unsupported.push('xlsx:table-auto-expand');
    }
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
      const result = await writeResult(document, options as InterchangeWriteOptions);
      if (result.diagnostics.length > 0) {
        const features = result.diagnostics.flatMap(({ details }) =>
          details !== undefined &&
          details !== null &&
          typeof details === 'object' &&
          'feature' in details &&
          typeof details.feature === 'string'
            ? [details.feature]
            : [],
        );
        throw new InterchangeError(
          'DOCUMENT_INVALID',
          `XLSX Blob export would omit unsupported semantics: ${features.join(', ')}`,
        );
      }
      return result.blob;
    },
    writeResult(document: SpreadsheetDocument, options = {}) {
      return writeResult(document, options as InterchangeWriteOptions);
    },
  });
}
