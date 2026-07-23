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
import { attributes, textContent } from './xml';

function positiveRepeat(value: string | undefined, label: string): number {
  if (value === undefined) return 1;
  const repeat = Number(value);
  if (!Number.isSafeInteger(repeat) || repeat < 1) {
    throw new InterchangeError('MALFORMED_WORKBOOK', `Invalid ${label} repeat`);
  }
  return repeat;
}

function odsCellInput(
  cellAttributes: Readonly<Record<string, string>>,
  body: string,
): CellInput | undefined {
  const formula = cellAttributes['table:formula'];
  if (formula !== undefined) {
    const normalized = formula.startsWith('of:') ? formula.slice(3) : formula;
    return { type: 'formula', source: normalized.startsWith('=') ? normalized : `=${normalized}` };
  }
  switch (cellAttributes['office:value-type']) {
    case undefined:
      return undefined;
    case 'string':
      return {
        type: 'string',
        value: [...body.matchAll(/<text:p\b[^>]*>([\s\S]*?)<\/text:p>/gi)]
          .map((match) => textContent(match[1]!))
          .join('\n'),
      };
    case 'float':
    case 'currency':
    case 'percentage': {
      const value = Number(cellAttributes['office:value']);
      if (!Number.isFinite(value)) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS numeric cell value is invalid');
      }
      return { type: 'number', value };
    }
    case 'boolean': {
      const value = cellAttributes['office:boolean-value'];
      if (value !== 'true' && value !== 'false') {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS boolean cell value is invalid');
      }
      return { type: 'boolean', value: value === 'true' };
    }
    case 'date':
    case 'time':
      return {
        type: 'string',
        value: cellAttributes['office:date-value'] ?? cellAttributes['office:time-value'] ?? '',
      };
    default:
      return {
        type: 'string',
        value: textContent(body),
      };
  }
}

function parseOdsSheet(
  name: string,
  body: string,
  limits: ReturnType<typeof resolveLimits>,
  unsupported: string[],
  state: { cells: number },
  signal?: AbortSignal,
): ImportedSheet {
  const cells: ImportedSheet['cells'][number][] = [];
  let rowIndex = 0;
  for (const rowMatch of body.matchAll(
    /<table:table-row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:table-row>)/gi,
  )) {
    throwIfAborted(signal);
    const rowRepeat = positiveRepeat(attributes(rowMatch[1]!)['table:number-rows-repeated'], 'row');
    if (rowIndex + rowRepeat > limits.maxRows) {
      throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'ODS row limit exceeded');
    }
    const parsedRow: { readonly column: number; readonly input: CellInput }[] = [];
    let column = 0;
    for (const cellMatch of (rowMatch[2] ?? '').matchAll(
      /<table:(?:table-cell|covered-table-cell)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:(?:table-cell|covered-table-cell)>)/gi,
    )) {
      const cellAttributes = attributes(cellMatch[1]!);
      const columnRepeat = positiveRepeat(
        cellAttributes['table:number-columns-repeated'],
        'column',
      );
      if (column + columnRepeat > limits.maxColumns) {
        throw new InterchangeError('COLUMN_LIMIT_EXCEEDED', 'ODS column limit exceeded');
      }
      if (
        cellAttributes['table:number-columns-spanned'] !== undefined ||
        cellAttributes['table:number-rows-spanned'] !== undefined
      ) {
        unsupported.push('ods:merged-cells');
      }
      const input = odsCellInput(cellAttributes, cellMatch[2] ?? '');
      if (input) {
        for (let offset = 0; offset < columnRepeat; offset += 1) {
          parsedRow.push({ column: column + offset, input });
        }
      }
      column += columnRepeat;
    }
    state.cells += parsedRow.length * rowRepeat;
    if (state.cells > limits.maxCells) {
      throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'ODS cell limit exceeded');
    }
    for (let repeatedRow = 0; repeatedRow < rowRepeat; repeatedRow += 1) {
      for (const cell of parsedRow) {
        cells.push({ row: rowIndex + repeatedRow, column: cell.column, input: cell.input });
      }
    }
    rowIndex += rowRepeat;
  }
  return { name, cells };
}

/** Creates an atomic ODS reader with ZIP and XML resource limits. */
export function createOdsReader(configuredLimits: InterchangeLimits = {}): WorkbookReader {
  const limits = resolveLimits(configuredLimits);
  return Object.freeze({
    format: 'ods',
    async read(
      input: InterchangeInput,
      options: InterchangeReadOptions = {},
    ): Promise<WorkbookImportResult> {
      const entries = await readArchive(input, limits, options.signal);
      if (
        Object.keys(entries).some((name) => /(?:^|\/)(?:scripts|basic|macros)(?:\/|$)/i.test(name))
      ) {
        throw new InterchangeError('ACTIVE_CONTENT_REJECTED', 'ODS active content is disabled');
      }
      const content = archiveXml(entries, 'content.xml', limits);
      if (/\bxlink:href\s*=\s*["'](?:[a-z][\w+.-]*:|\/\/)/i.test(content)) {
        throw new InterchangeError('EXTERNAL_RESOURCE_REJECTED', 'External ODS links are disabled');
      }
      const unsupported: string[] = [];
      if (/<table:shapes\b|<draw:/i.test(content)) unsupported.push('ods:drawing-objects');
      if (/<table:database-ranges\b/i.test(content)) unsupported.push('ods:database-ranges');
      if (/<table:named-expressions\b/i.test(content)) unsupported.push('ods:named-expressions');
      if (/<table:content-validations\b/i.test(content)) unsupported.push('ods:data-validation');
      if (/<table:data-pilot-tables\b/i.test(content)) unsupported.push('ods:pivot-tables');
      if (/<office:forms\b/i.test(content)) unsupported.push('ods:forms');
      if (/<office:annotation\b/i.test(content)) unsupported.push('ods:comments');
      if (/<table:filter\b/i.test(content)) unsupported.push('ods:filters');
      const sheets: ImportedSheet[] = [];
      const state = { cells: 0 };
      for (const match of content.matchAll(/<table:table\b([^>]*?)>([\s\S]*?)<\/table:table>/gi)) {
        const tableAttributes = attributes(match[1]!);
        if (!tableAttributes['table:name']) {
          throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS sheet name is missing');
        }
        sheets.push(
          parseOdsSheet(
            tableAttributes['table:name'],
            match[2]!,
            limits,
            unsupported,
            state,
            options.signal,
          ),
        );
      }
      if (sheets.length === 0) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS contains no worksheets');
      }
      throwIfAborted(options.signal);
      return importResult('ods', buildDocument(sheets), [], unsupported);
    },
  });
}
