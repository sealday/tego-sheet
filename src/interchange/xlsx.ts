import type { CellInput } from '../document';
import { archiveXml, readArchive } from './archive';
import {
  importResult,
  InterchangeError,
  resolveLimits,
  throwIfAborted,
  type InterchangeInput,
  type InterchangeLimits,
  type InterchangeReadOptions,
  type WorkbookImportResult,
  type WorkbookReader,
} from './contracts';
import { buildDocument, type ImportedSheet } from './document-builder';
import { attributes, decodeXml, textContent } from './xml';

function columnIndex(reference: string): number {
  const letters = /^([A-Z]+)\d+$/i.exec(reference)?.[1];
  if (!letters)
    throw new InterchangeError('MALFORMED_WORKBOOK', `Invalid cell reference: ${reference}`);
  let result = 0;
  for (const character of letters.toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result - 1;
}

function rowIndex(reference: string): number {
  const row = Number.parseInt(/\d+$/.exec(reference)?.[0] ?? '', 10) - 1;
  if (!Number.isSafeInteger(row) || row < 0) {
    throw new InterchangeError('MALFORMED_WORKBOOK', `Invalid cell reference: ${reference}`);
  }
  return row;
}

function relationshipTargets(xml: string): ReadonlyMap<string, string> {
  const targets = new Map<string, string>();
  for (const match of xml.matchAll(/<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/gi)) {
    const attribute = attributes(match[1]!);
    if (
      attribute.TargetMode?.toLowerCase() === 'external' ||
      /^[a-z][\w+.-]*:/i.test(attribute.Target ?? '')
    ) {
      throw new InterchangeError(
        'EXTERNAL_RESOURCE_REJECTED',
        'External workbook relationships are disabled',
      );
    }
    if (attribute.Id && attribute.Target) targets.set(attribute.Id, attribute.Target);
  }
  return targets;
}

function normalizedWorksheetTarget(target: string): string {
  const segments = `xl/${target}`.split('/');
  const normalized: string[] = [];
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (normalized.length === 0) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'Worksheet target escapes package root');
      }
      normalized.pop();
    } else {
      normalized.push(segment);
    }
  }
  const result = normalized.join('/');
  if (!result.startsWith('xl/worksheets/')) {
    throw new InterchangeError('MALFORMED_WORKBOOK', 'Worksheet target is outside worksheet parts');
  }
  return result;
}

function sharedStrings(xml: string | undefined): readonly string[] {
  if (!xml) return [];
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)].map((match) =>
    [...match[1]!.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
      .map((text) => textContent(text[1]!))
      .join(''),
  );
}

function parseCell(
  cellAttributes: Readonly<Record<string, string>>,
  body: string,
  strings: readonly string[],
  unsupported: string[],
): CellInput | undefined {
  const formula = /<f\b[^>]*>([\s\S]*?)<\/f>/i.exec(body);
  if (formula) return { type: 'formula', source: `=${textContent(formula[1]!)}` };
  const type = cellAttributes.t;
  if (type === 'inlineStr') {
    return {
      type: 'string',
      value: [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
        .map((match) => textContent(match[1]!))
        .join(''),
    };
  }
  const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/i.exec(body)?.[1];
  if (raw === undefined) return undefined;
  const value = decodeXml(raw.trim());
  if (type === 's') {
    const index = Number.parseInt(value, 10);
    if (!Number.isSafeInteger(index) || strings[index] === undefined) {
      throw new InterchangeError('MALFORMED_WORKBOOK', 'Shared string index is invalid');
    }
    return { type: 'string', value: strings[index] };
  }
  if (type === 'str') return { type: 'string', value };
  if (type === 'b') {
    if (value !== '0' && value !== '1') {
      throw new InterchangeError('MALFORMED_WORKBOOK', 'Boolean cell value is invalid');
    }
    return { type: 'boolean', value: value === '1' };
  }
  if (type === 'e') {
    unsupported.push('xlsx:error-cells');
    return { type: 'string', value };
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new InterchangeError('MALFORMED_WORKBOOK', 'Numeric cell value is invalid');
  }
  return { type: 'number', value: number };
}

function parseWorksheet(
  name: string,
  xml: string,
  strings: readonly string[],
  limits: ReturnType<typeof resolveLimits>,
  unsupported: string[],
  signal?: AbortSignal,
): ImportedSheet {
  const cells: ImportedSheet['cells'][number][] = [];
  const pattern = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/gi;
  for (const match of xml.matchAll(pattern)) {
    if ((cells.length & 0x3fff) === 0) throwIfAborted(signal);
    const cellAttributes = attributes(match[1]!);
    const reference = cellAttributes.r;
    if (!reference) throw new InterchangeError('MALFORMED_WORKBOOK', 'Cell reference is missing');
    const row = rowIndex(reference);
    const column = columnIndex(reference);
    if (row >= limits.maxRows) {
      throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'XLSX row limit exceeded');
    }
    if (column >= limits.maxColumns) {
      throw new InterchangeError('COLUMN_LIMIT_EXCEEDED', 'XLSX column limit exceeded');
    }
    const input = parseCell(cellAttributes, match[2] ?? '', strings, unsupported);
    if (input) {
      cells.push({ row, column, input });
      if (cells.length > limits.maxCells) {
        throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'XLSX cell limit exceeded');
      }
    }
  }
  if (/<mergeCells\b/i.test(xml)) unsupported.push('xlsx:merged-cells');
  if (/<conditionalFormatting\b/i.test(xml)) unsupported.push('xlsx:conditional-formatting');
  if (/<dataValidations\b/i.test(xml)) unsupported.push('xlsx:data-validation');
  return { name, cells };
}

export function createXlsxReader(configuredLimits: InterchangeLimits = {}): WorkbookReader {
  const limits = resolveLimits(configuredLimits);
  return Object.freeze({
    format: 'xlsx',
    async read(
      input: InterchangeInput,
      options: InterchangeReadOptions = {},
    ): Promise<WorkbookImportResult> {
      const entries = await readArchive(input, limits, options.signal);
      const entryNames = Object.keys(entries);
      if (
        entryNames.some((name) =>
          /(?:^|\/)(?:vbaProject\.bin|macrosheets?|xl4macros|activeX)(?:\/|$)/i.test(name),
        )
      ) {
        throw new InterchangeError('ACTIVE_CONTENT_REJECTED', 'XLSX active content is disabled');
      }
      const contentTypes = archiveXml(entries, '[Content_Types].xml', limits);
      if (/macroEnabled|vbaProject|activeX/i.test(contentTypes)) {
        throw new InterchangeError(
          'ACTIVE_CONTENT_REJECTED',
          'Macro-enabled XLSX content types are disabled',
        );
      }
      for (const name of entryNames.filter((entry) => /\.rels$/i.test(entry))) {
        relationshipTargets(archiveXml(entries, name, limits));
      }
      const workbook = archiveXml(entries, 'xl/workbook.xml', limits);
      const relationships = relationshipTargets(
        archiveXml(entries, 'xl/_rels/workbook.xml.rels', limits),
      );
      const shared =
        entries['xl/sharedStrings.xml'] === undefined
          ? undefined
          : archiveXml(entries, 'xl/sharedStrings.xml', limits);
      const strings = sharedStrings(shared);
      const unsupported: string[] = [];
      const sheets: ImportedSheet[] = [];
      let totalCells = 0;
      for (const match of workbook.matchAll(/<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/gi)) {
        throwIfAborted(options.signal);
        const sheetAttributes = attributes(match[1]!);
        const relationshipId = sheetAttributes['r:id'];
        const target = relationshipId ? relationships.get(relationshipId) : undefined;
        if (!sheetAttributes.name || !target) {
          throw new InterchangeError('MALFORMED_WORKBOOK', 'Worksheet relationship is missing');
        }
        const worksheetName = normalizedWorksheetTarget(target);
        const sheet = parseWorksheet(
          sheetAttributes.name,
          archiveXml(entries, worksheetName, limits),
          strings,
          limits,
          unsupported,
          options.signal,
        );
        totalCells += sheet.cells.length;
        if (totalCells > limits.maxCells) {
          throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'XLSX cell limit exceeded');
        }
        sheets.push(sheet);
      }
      if (sheets.length === 0) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'Workbook contains no worksheets');
      }
      throwIfAborted(options.signal);
      return importResult('xlsx', buildDocument(sheets), [], unsupported);
    },
  });
}
