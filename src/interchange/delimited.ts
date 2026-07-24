import type { CellInput, SpreadsheetDocument } from '../document';
import {
  exportResult,
  importResult,
  inputBytes,
  InterchangeError,
  resolveLimits,
  throwIfAborted,
  type DelimitedWriteOptions,
  type InterchangeInput,
  type InterchangeLimits,
  type InterchangeReadOptions,
  type WorkbookImportResult,
  type WorkbookReader,
  type WorkbookWriter,
} from './contracts';
import { buildDocument, type ImportedSheet } from './document-builder';

function utf8Bytes(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function inferredInput(value: string): CellInput | undefined {
  if (value === '') return undefined;
  if (/^true$/i.test(value)) return { type: 'boolean', value: true };
  if (/^false$/i.test(value)) return { type: 'boolean', value: false };
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(value)) {
    const number = Number(value);
    if (Number.isFinite(number)) return { type: 'number', value: number };
  }
  return { type: 'string', value };
}

function parseRows(
  text: string,
  delimiter: ',' | '\t',
  limits: ReturnType<typeof resolveLimits>,
  signal?: AbortSignal,
): ImportedSheet {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let afterQuote = false;
  let logicalCells = 0;

  const appendField = (): void => {
    if (utf8Bytes(field) > limits.maxFieldBytes) {
      throw new InterchangeError('FIELD_LIMIT_EXCEEDED', 'Delimited field byte limit exceeded');
    }
    row.push(field);
    field = '';
    afterQuote = false;
    logicalCells += 1;
    if (logicalCells > limits.maxCells) {
      throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'Delimited cell limit exceeded');
    }
    if (row.length > limits.maxColumns) {
      throw new InterchangeError('COLUMN_LIMIT_EXCEEDED', 'Delimited column limit exceeded');
    }
  };
  const appendRow = (): void => {
    appendField();
    rows.push(row);
    row = [];
    if (rows.length > limits.maxRows) {
      throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'Delimited row limit exceeded');
    }
  };

  for (let index = 0; index < text.length; index += 1) {
    if ((index & 0x3fff) === 0) throwIfAborted(signal);
    const character = text[index]!;
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"';
          index += 1;
        } else {
          quoted = false;
          afterQuote = true;
        }
      } else {
        field += character;
      }
      continue;
    }
    if (afterQuote && character !== delimiter && character !== '\r' && character !== '\n') {
      throw new InterchangeError(
        'MALFORMED_DELIMITED_TEXT',
        'Unexpected character after closing quote',
      );
    }
    if (character === '"' && field === '' && !afterQuote) {
      quoted = true;
    } else if (character === '"') {
      throw new InterchangeError(
        'MALFORMED_DELIMITED_TEXT',
        'Quote in an unquoted field must be escaped',
      );
    } else if (character === delimiter) {
      appendField();
    } else if (character === '\r' || character === '\n') {
      appendRow();
      if (character === '\r' && text[index + 1] === '\n') index += 1;
    } else {
      field += character;
    }
  }
  if (quoted) {
    throw new InterchangeError('MALFORMED_DELIMITED_TEXT', 'Delimited input has an open quote');
  }
  const endsInLineBreak = text.endsWith('\n') || text.endsWith('\r');
  if (!endsInLineBreak && (field !== '' || row.length > 0 || text.length > 0)) appendRow();

  const cells: ImportedSheet['cells'][number][] = [];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let column = 0; column < rows[rowIndex]!.length; column += 1) {
      const input = inferredInput(rows[rowIndex]![column]!);
      if (input) cells.push({ row: rowIndex, column, input });
    }
  }
  return { name: 'Sheet 1', cells };
}

function createDelimitedReader(
  format: 'csv' | 'tsv',
  delimiter: ',' | '\t',
  configuredLimits: InterchangeLimits = {},
): WorkbookReader {
  const limits = resolveLimits(configuredLimits);
  return Object.freeze({
    format,
    async read(
      input: InterchangeInput,
      options: InterchangeReadOptions = {},
    ): Promise<WorkbookImportResult> {
      const bytes = await inputBytes(input, options.signal);
      if (bytes.byteLength > limits.maxTextBytes) {
        throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'Delimited input byte limit exceeded');
      }
      let text: string;
      try {
        text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
      } catch (cause) {
        throw new InterchangeError('MALFORMED_DELIMITED_TEXT', 'Delimited input is not UTF-8', {
          cause,
        });
      }
      const sheet = parseRows(text, delimiter, limits, options.signal);
      throwIfAborted(options.signal);
      return importResult(format, buildDocument([sheet]));
    },
  });
}

function cellText(input: CellInput): string {
  switch (input.type) {
    case 'blank':
      return '';
    case 'boolean':
      return input.value ? 'TRUE' : 'FALSE';
    case 'number':
      return String(input.value);
    case 'formula':
      return input.source;
    case 'string':
      return input.value;
    case 'custom':
      return JSON.stringify(input.value);
  }
}

function protectFormula(value: string): string {
  return /^[\t\r\n ]*[=+\-@]/.test(value) ? `'${value}` : value;
}

function quoted(value: string, delimiter: string): string {
  return value.includes(delimiter) || /["\r\n]/.test(value)
    ? `"${value.replaceAll('"', '""')}"`
    : value;
}

function createDelimitedWriter(
  format: 'csv' | 'tsv',
  defaultDelimiter: ',' | '\t',
  configuredLimits: InterchangeLimits = {},
): WorkbookWriter {
  const limits = resolveLimits(configuredLimits);
  const write = async (
    document: SpreadsheetDocument,
    options: DelimitedWriteOptions = {},
  ): Promise<Blob> => {
    throwIfAborted(options.signal);
    const delimiter = options.delimiter ?? defaultDelimiter;
    if (delimiter !== ',' && delimiter !== '\t') {
      throw new TypeError('delimiter must be a comma or tab');
    }
    const lineEnding = options.lineEnding ?? '\n';
    const sheet = document.workbook.sheets[0];
    const maxRow = sheet?.cells.reduce((maximum, cell) => Math.max(maximum, cell.row), -1) ?? -1;
    const maxColumn =
      sheet?.cells.reduce((maximum, cell) => Math.max(maximum, cell.column), -1) ?? -1;
    if (maxRow + 1 > limits.maxRows) {
      throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'Delimited row limit exceeded');
    }
    if (maxColumn + 1 > limits.maxColumns) {
      throw new InterchangeError('COLUMN_LIMIT_EXCEEDED', 'Delimited column limit exceeded');
    }
    const byCoordinate = new Map(
      sheet?.cells.map((cell) => [`${cell.row}:${cell.column}`, cell.cell.input]) ?? [],
    );
    const rows: string[] = [];
    for (let row = 0; row <= maxRow; row += 1) {
      throwIfAborted(options.signal);
      const fields: string[] = [];
      for (let column = 0; column <= maxColumn; column += 1) {
        const input = byCoordinate.get(`${row}:${column}`);
        let value = input ? cellText(input) : '';
        if (options.formulaInjectionProtection !== false && input?.type === 'string') {
          value = protectFormula(value);
        }
        fields.push(quoted(value, delimiter));
      }
      rows.push(fields.join(delimiter));
    }
    const text = rows.length === 0 ? '' : `${rows.join(lineEnding)}${lineEnding}`;
    const bytes = utf8Bytes(text);
    const maxOutputBytes = options.maxOutputBytes ?? limits.maxOutputBytes;
    if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 0) {
      throw new TypeError('maxOutputBytes must be a non-negative safe integer');
    }
    if (bytes > maxOutputBytes) {
      throw new InterchangeError('OUTPUT_LIMIT_EXCEEDED', 'Delimited output byte limit exceeded');
    }
    throwIfAborted(options.signal);
    return new Blob([text], {
      type: delimiter === ',' ? 'text/csv' : 'text/tab-separated-values',
    });
  };
  return Object.freeze({
    format,
    write,
    async writeResult(document: SpreadsheetDocument, options = {}) {
      return exportResult(format, await write(document, options as DelimitedWriteOptions));
    },
  });
}

/** Creates an isolated bounded CSV reader. */
export function createCsvReader(limits?: InterchangeLimits): WorkbookReader {
  return createDelimitedReader('csv', ',', limits);
}

/** Creates an isolated bounded TSV reader. */
export function createTsvReader(limits?: InterchangeLimits): WorkbookReader {
  return createDelimitedReader('tsv', '\t', limits);
}

/** Creates an isolated bounded CSV writer. */
export function createCsvWriter(limits?: InterchangeLimits): WorkbookWriter {
  return createDelimitedWriter('csv', ',', limits);
}

/** Creates an isolated bounded TSV writer. */
export function createTsvWriter(limits?: InterchangeLimits): WorkbookWriter {
  return createDelimitedWriter('tsv', '\t', limits);
}
