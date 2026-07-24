import { strToU8, zipSync } from 'fflate';
import type {
  Cell,
  CellInput,
  DocumentSheetId,
  JsonValue,
  Sheet,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../document';
import type { TemplatePrintProfile } from '../template';
import { archiveXml, readArchive } from './archive';
import {
  exportResult,
  importResult,
  InterchangeError,
  resolveLimits,
  throwIfAborted,
  type InterchangeInput,
  type InterchangeLimits,
  type InterchangeReadOptions,
  type InterchangeWriteOptions,
  type WorkbookImportResult,
  type WorkbookExportResult,
  type WorkbookReader,
  type WorkbookWriter,
} from './contracts';
import { buildDocument, type ImportedSheet } from './document-builder';
import { attributes, textContent } from './xml';

const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';
const FIXED_ZIP_DATE = new Date(1980, 0, 1, 0, 0, 0, 0);

function xml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function odsCell(
  cell: Cell | undefined,
  styleNames: ReadonlyMap<string, string>,
  validationNames: ReadonlyMap<string, string>,
): string {
  if (cell === undefined || cell.input.type === 'blank') return '<table:table-cell/>';
  const styleName = cell.styleId === undefined ? undefined : styleNames.get(cell.styleId);
  const validationName =
    cell.validationId === undefined ? undefined : validationNames.get(cell.validationId);
  const semantic =
    `${styleName === undefined ? '' : ` table:style-name="${xml(styleName)}"`}` +
    `${validationName === undefined ? '' : ` table:content-validation-name="${xml(validationName)}"`}`;
  const input = cell.input;
  if (input.type === 'string') {
    return `<table:table-cell${semantic} office:value-type="string"><text:p>${xml(input.value)}</text:p></table:table-cell>`;
  }
  if (input.type === 'number') {
    return `<table:table-cell${semantic} office:value-type="float" office:value="${input.value}"/>`;
  }
  if (input.type === 'boolean') {
    return `<table:table-cell${semantic} office:value-type="boolean" office:boolean-value="${input.value ? 'true' : 'false'}"/>`;
  }
  if (input.type === 'formula') {
    return `<table:table-cell${semantic} table:formula="of:${xml(input.source)}"/>`;
  }
  throw new InterchangeError(
    'DOCUMENT_INVALID',
    `ODS cannot serialize custom cell type ${input.cellType}`,
  );
}

function odsSheet(
  sheet: Sheet,
  limits: ReturnType<typeof resolveLimits>,
  styleNames: ReadonlyMap<string, string>,
  validationNames: ReadonlyMap<string, string>,
  hasPrintProfile: boolean,
): string {
  let maxRow = Math.max(-1, (sheet.rowCount ?? 0) - 1);
  let maxColumn = Math.max(-1, (sheet.columnCount ?? 0) - 1);
  for (const entry of sheet.cells) {
    maxRow = Math.max(maxRow, entry.row);
    maxColumn = Math.max(maxColumn, entry.column);
  }
  for (const merge of sheet.merges) {
    maxRow = Math.max(maxRow, merge.end.row);
    maxColumn = Math.max(maxColumn, merge.end.column);
  }
  if (maxRow + 1 > limits.maxRows) {
    throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'ODS row limit exceeded');
  }
  if (maxColumn + 1 > limits.maxColumns) {
    throw new InterchangeError('COLUMN_LIMIT_EXCEEDED', 'ODS column limit exceeded');
  }
  const byCoordinate = new Map(
    sheet.cells.map((entry) => [`${entry.row}:${entry.column}`, entry.cell]),
  );
  const mergeByAnchor = new Map(
    sheet.merges.map((merge) => [`${merge.start.row}:${merge.start.column}`, merge]),
  );
  const covered = new Set<string>();
  for (const merge of sheet.merges) {
    for (let row = merge.start.row; row <= merge.end.row; row += 1) {
      for (let column = merge.start.column; column <= merge.end.column; column += 1) {
        if (row !== merge.start.row || column !== merge.start.column) {
          covered.add(`${row}:${column}`);
        }
      }
    }
  }
  const rows: string[] = [];
  for (let row = 0; row <= maxRow; row += 1) {
    const cells: string[] = [];
    for (let column = 0; column <= maxColumn; column += 1) {
      const key = `${row}:${column}`;
      if (covered.has(key)) {
        cells.push('<table:covered-table-cell/>');
        continue;
      }
      const merge = mergeByAnchor.get(key);
      const serialized = odsCell(byCoordinate.get(key), styleNames, validationNames);
      cells.push(
        merge === undefined
          ? serialized
          : serialized.replace(
              '<table:table-cell',
              `<table:table-cell table:number-columns-spanned="${merge.end.column - merge.start.column + 1}" table:number-rows-spanned="${merge.end.row - merge.start.row + 1}"`,
            ),
      );
    }
    rows.push(`<table:table-row>${cells.join('')}</table:table-row>`);
  }
  return `<table:table table:name="${xml(sheet.name)}"${hasPrintProfile ? ' table:style-name="tego-print-table"' : ''}>${rows.join('')}</table:table>`;
}

function jsonRecord(value: JsonValue): Readonly<Record<string, JsonValue>> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Readonly<Record<string, JsonValue>>)
    : {};
}

function odsStyles(document: SpreadsheetDocument): {
  readonly names: ReadonlyMap<string, string>;
  readonly xml: string;
} {
  const names = new Map(
    document.workbook.styles.map((entry, index) => [entry.id, `ce${index + 1}`]),
  );
  const numberStyles: string[] = [];
  const styles = document.workbook.styles.map((entry, index) => {
    const value = jsonRecord(entry.value);
    const name = names.get(entry.id)!;
    let dataStyle = '';
    if (typeof value.numberFormat === 'string' && /^0(?:\.(0+))?$/.test(value.numberFormat)) {
      const decimals = /^0(?:\.(0+))?$/.exec(value.numberFormat)?.[1]?.length ?? 0;
      const numberName = `number${index + 1}`;
      numberStyles.push(
        `<number:number-style style:name="${numberName}"><number:number number:min-integer-digits="1" number:decimal-places="${decimals}"/></number:number-style>`,
      );
      dataStyle = ` style:data-style-name="${numberName}"`;
    }
    const textProperties =
      typeof value.color === 'string' ||
      typeof value.fontFamily === 'string' ||
      typeof value.fontSize === 'number' ||
      value.bold === true ||
      value.italic === true
        ? `<style:text-properties${typeof value.color === 'string' ? ` fo:color="${xml(value.color)}"` : ''}${typeof value.fontFamily === 'string' ? ` style:font-name="${xml(value.fontFamily)}"` : ''}${typeof value.fontSize === 'number' ? ` fo:font-size="${value.fontSize}pt"` : ''}${value.bold === true ? ' fo:font-weight="bold"' : ''}${value.italic === true ? ' fo:font-style="italic"' : ''}/>`
        : '';
    const cellProperties =
      typeof value.backgroundColor === 'string' || typeof value.verticalAlign === 'string'
        ? `<style:table-cell-properties${typeof value.backgroundColor === 'string' ? ` fo:background-color="${xml(value.backgroundColor)}"` : ''}${typeof value.verticalAlign === 'string' ? ` style:vertical-align="${value.verticalAlign}"` : ''}/>`
        : '';
    const paragraph =
      typeof value.horizontalAlign === 'string'
        ? `<style:paragraph-properties fo:text-align="${value.horizontalAlign}"/>`
        : '';
    return `<style:style style:name="${name}" style:family="table-cell"${dataStyle}>${textProperties}${cellProperties}${paragraph}</style:style>`;
  });
  return { names, xml: `${numberStyles.join('')}${styles.join('')}` };
}

function odsValidations(document: SpreadsheetDocument): {
  readonly names: ReadonlyMap<string, string>;
  readonly xml: string;
  readonly unsupported: readonly string[];
} {
  const names = new Map<string, string>();
  const unsupported: string[] = [];
  const definitions = document.workbook.validations.flatMap((entry, index) => {
    const value = jsonRecord(entry.value);
    if (value.type !== 'list' || typeof value.formula1 !== 'string') {
      unsupported.push(`ods:validation-${String(value.type ?? 'unknown')}`);
      return [];
    }
    const validationName = `validation${index + 1}`;
    names.set(entry.id, validationName);
    const list =
      value.formula1.startsWith('"') && value.formula1.endsWith('"')
        ? value.formula1.slice(1, -1).split(',')
        : [value.formula1];
    const condition = `of:cell-content-is-in-list(${list.map((item) => `"${item.replaceAll('"', '""')}"`).join(';')})`;
    return [
      `<table:content-validation table:name="${validationName}" table:condition="${xml(condition)}" table:allow-empty-cell="${value.allowBlank === true ? 'true' : 'false'}"/>`,
    ];
  });
  return {
    names,
    xml:
      definitions.length === 0
        ? ''
        : `<table:content-validations>${definitions.join('')}</table:content-validations>`,
    unsupported,
  };
}

function odsContent(
  document: SpreadsheetDocument,
  limits: ReturnType<typeof resolveLimits>,
  styles: ReturnType<typeof odsStyles>,
  validations: ReturnType<typeof odsValidations>,
): string {
  if (
    document.workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0) > limits.maxCells
  ) {
    throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'ODS cell limit exceeded');
  }
  const profile = document.templates.flatMap((template) => template.printProfiles)[0];
  const printOptions =
    profile === undefined
      ? ''
      : [profile.showGridlines ? 'grid' : '', profile.showHeadings ? 'headers' : '']
          .filter(Boolean)
          .join(' ');
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">' +
    `<office:automatic-styles>${styles.xml}<style:style style:name="tego-print-table" style:family="table" style:master-page-name="tego-master"><style:table-properties table:print="${printOptions}"/></style:style></office:automatic-styles>` +
    `<office:body><office:spreadsheet>${validations.xml}${document.workbook.sheets
      .map((sheet) =>
        odsSheet(sheet, limits, styles.names, validations.names, profile !== undefined),
      )
      .join('')}</office:spreadsheet></office:body></office:document-content>`
  );
}

function odsPrintStyles(document: SpreadsheetDocument): string {
  const profile = document.templates.flatMap((template) => template.printProfiles)[0];
  const page = profile?.page;
  const paper = page?.paper ?? { type: 'A4' as const };
  const dimensions =
    paper.type === 'custom'
      ? { width: paper.width / 96, height: paper.height / 96 }
      : paper.type === 'A5'
        ? { width: 5.83, height: 8.27 }
        : paper.type === 'Letter'
          ? { width: 8.5, height: 11 }
          : { width: 8.27, height: 11.69 };
  const margins = page?.margins ?? { top: 48, right: 48, bottom: 48, left: 48 };
  const bandText = (band: TemplatePrintProfile['header'] | undefined): string =>
    band?.center ?? band?.left ?? band?.right ?? '';
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">' +
    `<office:automatic-styles><style:page-layout style:name="tego-page"><style:page-layout-properties fo:page-width="${dimensions.width}in" fo:page-height="${dimensions.height}in" style:print-orientation="${page?.orientation ?? 'portrait'}" fo:margin-top="${margins.top / 96}in" fo:margin-right="${margins.right / 96}in" fo:margin-bottom="${margins.bottom / 96}in" fo:margin-left="${margins.left / 96}in"/></style:page-layout></office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="tego-master" style:page-layout-name="tego-page"><style:header><text:p>${xml(bandText(profile?.header))}</text:p></style:header><style:footer><text:p>${xml(bandText(profile?.footer))}</text:p></style:footer></style:master-page></office:master-styles>` +
    '</office:document-styles>'
  );
}

function positiveRepeat(value: string | undefined, label: string): number {
  if (value === undefined) return 1;
  const repeat = Number(value);
  if (!Number.isSafeInteger(repeat) || repeat < 1) {
    throw new InterchangeError('MALFORMED_WORKBOOK', `Invalid ${label} repeat`);
  }
  return repeat;
}

interface ParsedOdsRegistries {
  readonly styles: readonly { readonly id: string; readonly value: JsonValue }[];
  readonly styleIds: ReadonlyMap<string, string>;
  readonly validations: readonly { readonly id: string; readonly value: JsonValue }[];
  readonly validationIds: ReadonlyMap<string, string>;
  readonly unsupported: readonly string[];
}

function dimensionInches(value: string | undefined): number {
  if (value === undefined) return 0;
  const number = Number.parseFloat(value);
  if (!Number.isFinite(number)) return 0;
  if (value.endsWith('cm')) return number / 2.54;
  if (value.endsWith('mm')) return number / 25.4;
  if (value.endsWith('pt')) return number / 72;
  return number;
}

function parseOdsPrintProfile(
  content: string,
  stylesXml: string | undefined,
  sheetIds: readonly string[],
): SpreadsheetDocumentInput['templates'][number]['printProfiles'][number] | undefined {
  if (stylesXml === undefined || !/<style:page-layout-properties\b/i.test(stylesXml)) {
    return undefined;
  }
  const page = attributes(
    /<style:page-layout-properties\b([^>]*?)(?:\/>|>)/i.exec(stylesXml)?.[1] ?? '',
  );
  const table = attributes(/<style:table-properties\b([^>]*?)(?:\/>|>)/i.exec(content)?.[1] ?? '');
  const width = dimensionInches(page['fo:page-width']);
  const height = dimensionInches(page['fo:page-height']);
  const paper =
    Math.abs(width - 8.27) < 0.05 && Math.abs(height - 11.69) < 0.05
      ? ({ type: 'A4' } as const)
      : Math.abs(width - 5.83) < 0.05 && Math.abs(height - 8.27) < 0.05
        ? ({ type: 'A5' } as const)
        : Math.abs(width - 8.5) < 0.05 && Math.abs(height - 11) < 0.05
          ? ({ type: 'Letter' } as const)
          : ({ type: 'custom', width: width * 96, height: height * 96 } as const);
  const header = textContent(
    /<style:header\b[^>]*>([\s\S]*?)<\/style:header>/i.exec(stylesXml)?.[1] ?? '',
  );
  const footer = textContent(
    /<style:footer\b[^>]*>([\s\S]*?)<\/style:footer>/i.exec(stylesXml)?.[1] ?? '',
  );
  const print = new Set((table['table:print'] ?? '').split(/\s+/));
  return {
    id: 'ods-imported-print',
    name: 'Imported print',
    targets: sheetIds.map((sheetId) => ({
      type: 'sheet' as const,
      sheetId: sheetId as DocumentSheetId,
    })),
    page: {
      paper,
      orientation: page['style:print-orientation'] === 'landscape' ? 'landscape' : 'portrait',
      margins: {
        top: dimensionInches(page['fo:margin-top']) * 96,
        right: dimensionInches(page['fo:margin-right']) * 96,
        bottom: dimensionInches(page['fo:margin-bottom']) * 96,
        left: dimensionInches(page['fo:margin-left']) * 96,
      },
      scale: { type: 'fixed', value: 1 },
    },
    manualBreaks: [],
    ...(header === '' ? {} : { header: { center: header } }),
    ...(footer === '' ? {} : { footer: { center: footer } }),
    showGridlines: print.has('grid'),
    showHeadings: print.has('headers'),
  };
}

function parseOdsRegistries(content: string): ParsedOdsRegistries {
  const numberFormats = new Map<string, string>();
  for (const match of content.matchAll(
    /<number:number-style\b([^>]*?)>([\s\S]*?)<\/number:number-style>/gi,
  )) {
    const name = attributes(match[1]!)['style:name'];
    const number = /<number:number\b([^>]*?)(?:\/>|>)/i.exec(match[2]!);
    if (name === undefined || number === null) continue;
    const decimals = Number(attributes(number[1]!)['number:decimal-places'] ?? 0);
    if (Number.isSafeInteger(decimals) && decimals >= 0) {
      numberFormats.set(name, decimals === 0 ? '0' : `0.${'0'.repeat(decimals)}`);
    }
  }
  const styles: { id: string; value: JsonValue }[] = [];
  const styleIds = new Map<string, string>();
  for (const match of content.matchAll(/<style:style\b([^>]*?)>([\s\S]*?)<\/style:style>/gi)) {
    const styleAttributes = attributes(match[1]!);
    if (
      styleAttributes['style:family'] !== 'table-cell' ||
      styleAttributes['style:name'] === undefined
    )
      continue;
    const value: Record<string, JsonValue> = {};
    const text = attributes(
      /<style:text-properties\b([^>]*?)(?:\/>|>)/i.exec(match[2]!)?.[1] ?? '',
    );
    const cell = attributes(
      /<style:table-cell-properties\b([^>]*?)(?:\/>|>)/i.exec(match[2]!)?.[1] ?? '',
    );
    const paragraph = attributes(
      /<style:paragraph-properties\b([^>]*?)(?:\/>|>)/i.exec(match[2]!)?.[1] ?? '',
    );
    if (text['fo:color'] !== undefined) value.color = text['fo:color'];
    if (text['style:font-name'] !== undefined) value.fontFamily = text['style:font-name'];
    if (text['fo:font-size'] !== undefined)
      value.fontSize = Number.parseFloat(text['fo:font-size']);
    if (text['fo:font-weight'] === 'bold') value.bold = true;
    if (text['fo:font-style'] === 'italic') value.italic = true;
    if (cell['fo:background-color'] !== undefined) {
      value.backgroundColor = cell['fo:background-color'];
    }
    if (cell['style:vertical-align'] !== undefined) {
      value.verticalAlign = cell['style:vertical-align'];
    }
    if (paragraph['fo:text-align'] !== undefined) {
      value.horizontalAlign = paragraph['fo:text-align'];
    }
    const numberFormat = numberFormats.get(styleAttributes['style:data-style-name'] ?? '');
    if (numberFormat !== undefined) value.numberFormat = numberFormat;
    const id = `ods-style-${styles.length + 1}`;
    styles.push({ id, value });
    styleIds.set(styleAttributes['style:name'], id);
  }
  const validations: { id: string; value: JsonValue }[] = [];
  const validationIds = new Map<string, string>();
  const unsupported: string[] = [];
  for (const match of content.matchAll(
    /<table:content-validation\b([^>]*?)(?:\/>|>[\s\S]*?<\/table:content-validation>)/gi,
  )) {
    const entry = attributes(match[1]!);
    if (entry['table:name'] === undefined || entry['table:condition'] === undefined) continue;
    const list = /^of:cell-content-is-in-list\((.*)\)$/i.exec(entry['table:condition']);
    if (list === null) {
      unsupported.push('ods:data-validation-condition');
      continue;
    }
    const values = list[1]!
      .split(';')
      .map((value) => value.trim().replace(/^"|"$/g, '').replaceAll('""', '"'));
    const id = `ods-validation-${validations.length + 1}`;
    validations.push({
      id,
      value: {
        type: 'list',
        formula1: `"${values.join(',')}"`,
        allowBlank: entry['table:allow-empty-cell'] === 'true',
      },
    });
    validationIds.set(entry['table:name'], id);
  }
  return { styles, styleIds, validations, validationIds, unsupported };
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
  id: string | undefined,
  name: string,
  body: string,
  limits: ReturnType<typeof resolveLimits>,
  unsupported: string[],
  registries: ParsedOdsRegistries,
  state: { cells: number },
  signal?: AbortSignal,
): ImportedSheet {
  const cells: ImportedSheet['cells'][number][] = [];
  const merges: NonNullable<ImportedSheet['merges']>[number][] = [];
  let rowIndex = 0;
  for (const rowMatch of body.matchAll(
    /<table:table-row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/table:table-row>)/gi,
  )) {
    throwIfAborted(signal);
    const rowRepeat = positiveRepeat(attributes(rowMatch[1]!)['table:number-rows-repeated'], 'row');
    if (rowIndex + rowRepeat > limits.maxRows) {
      throw new InterchangeError('ROW_LIMIT_EXCEEDED', 'ODS row limit exceeded');
    }
    const parsedRow: {
      readonly column: number;
      readonly input: CellInput;
      readonly styleId?: string;
      readonly validationId?: string;
    }[] = [];
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
      const columnsSpanned = positiveRepeat(
        cellAttributes['table:number-columns-spanned'],
        'column span',
      );
      const rowsSpanned = positiveRepeat(cellAttributes['table:number-rows-spanned'], 'row span');
      if (columnsSpanned > 1 || rowsSpanned > 1) {
        merges.push({
          start: { row: rowIndex, column },
          end: { row: rowIndex + rowsSpanned - 1, column: column + columnsSpanned - 1 },
        });
      }
      const input = odsCellInput(cellAttributes, cellMatch[2] ?? '');
      if (input) {
        for (let offset = 0; offset < columnRepeat; offset += 1) {
          parsedRow.push({
            column: column + offset,
            input,
            ...(cellAttributes['table:style-name'] === undefined
              ? {}
              : { styleId: registries.styleIds.get(cellAttributes['table:style-name']) }),
            ...(cellAttributes['table:content-validation-name'] === undefined
              ? {}
              : {
                  validationId: registries.validationIds.get(
                    cellAttributes['table:content-validation-name'],
                  ),
                }),
          });
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
        cells.push({
          row: rowIndex + repeatedRow,
          column: cell.column,
          input: cell.input,
          ...(cell.styleId === undefined ? {} : { styleId: cell.styleId }),
          ...(cell.validationId === undefined ? {} : { validationId: cell.validationId }),
        });
      }
    }
    rowIndex += rowRepeat;
  }
  return { ...(id === undefined ? {} : { id }), name, cells, merges };
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
      const stylesXml =
        entries['styles.xml'] === undefined ? undefined : archiveXml(entries, 'styles.xml', limits);
      if (/<office:scripts\b/i.test(content)) {
        throw new InterchangeError('ACTIVE_CONTENT_REJECTED', 'ODS active content is disabled');
      }
      if (/\bxlink:href\s*=\s*["'](?:[a-z][\w+.-]*:|\/\/)/i.test(`${content}${stylesXml ?? ''}`)) {
        throw new InterchangeError('EXTERNAL_RESOURCE_REJECTED', 'External ODS links are disabled');
      }
      const unsupported: string[] = [];
      if (/<table:shapes\b|<draw:/i.test(content)) unsupported.push('ods:drawing-objects');
      if (/<table:database-ranges\b/i.test(content)) unsupported.push('ods:database-ranges');
      if (/<table:named-expressions\b/i.test(content)) unsupported.push('ods:named-expressions');
      if (/<table:data-pilot-tables\b/i.test(content)) unsupported.push('ods:pivot-tables');
      if (/<office:forms\b/i.test(content)) unsupported.push('ods:forms');
      if (/<office:annotation\b/i.test(content)) unsupported.push('ods:comments');
      if (/<table:filter\b/i.test(content)) unsupported.push('ods:filters');
      const sheets: ImportedSheet[] = [];
      const registries = parseOdsRegistries(`${stylesXml ?? ''}${content}`);
      unsupported.push(...registries.unsupported);
      const state = { cells: 0 };
      for (const match of content.matchAll(/<table:table\b([^>]*?)>([\s\S]*?)<\/table:table>/gi)) {
        const tableAttributes = attributes(match[1]!);
        if (!tableAttributes['table:name']) {
          throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS sheet name is missing');
        }
        sheets.push(
          parseOdsSheet(
            `interchange-sheet-${sheets.length + 1}`,
            tableAttributes['table:name'],
            match[2]!,
            limits,
            unsupported,
            registries,
            state,
            options.signal,
          ),
        );
      }
      if (sheets.length === 0) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'ODS contains no worksheets');
      }
      throwIfAborted(options.signal);
      const profile = parseOdsPrintProfile(
        content,
        stylesXml,
        sheets.map((sheet) => sheet.id!),
      );
      return importResult(
        'ods',
        buildDocument(sheets, {
          styles: registries.styles,
          validations: registries.validations,
          templates:
            profile === undefined
              ? []
              : [
                  {
                    id: 'ods-imported-template',
                    name: 'Imported print',
                    bindings: [],
                    printProfiles: [profile],
                  },
                ],
        }),
        [],
        unsupported,
      );
    },
  });
}

/** Creates a deterministic bounded ODS writer for the supported semantic foundation. */
export function createOdsWriter(configuredLimits: InterchangeLimits = {}): WorkbookWriter {
  const limits = resolveLimits(configuredLimits);
  const writeResult = async (
    document: SpreadsheetDocument,
    options: InterchangeWriteOptions = {},
  ): Promise<WorkbookExportResult> => {
    throwIfAborted(options.signal);
    const styles = odsStyles(document);
    const validations = odsValidations(document);
    const content = odsContent(document, limits, styles, validations);
    const ordered: Parameters<typeof zipSync>[0] = {
      mimetype: [strToU8(ODS_MIME), { level: 0, mtime: FIXED_ZIP_DATE }],
      'META-INF/manifest.xml': [
        strToU8(
          '<?xml version="1.0" encoding="UTF-8"?>' +
            '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' +
            `<manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MIME}"/>` +
            '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
            '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
            '</manifest:manifest>',
        ),
        { level: 6, mtime: FIXED_ZIP_DATE },
      ],
      'content.xml': [strToU8(content), { level: 6, mtime: FIXED_ZIP_DATE }],
      'styles.xml': [strToU8(odsPrintStyles(document)), { level: 6, mtime: FIXED_ZIP_DATE }],
    };
    const bytes = zipSync(ordered, { level: 6, mtime: FIXED_ZIP_DATE });
    throwIfAborted(options.signal);
    const maxOutputBytes = options.maxOutputBytes ?? limits.maxOutputBytes;
    if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 0) {
      throw new TypeError('maxOutputBytes must be a non-negative safe integer');
    }
    if (bytes.byteLength > maxOutputBytes || bytes.byteLength > limits.maxPackageBytes) {
      throw new InterchangeError('OUTPUT_LIMIT_EXCEEDED', 'ODS output byte limit exceeded');
    }
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    const blob = new Blob([buffer], { type: ODS_MIME });
    const unsupported = [
      ...(document.workbook.sheets.some((sheet) => sheet.filter !== undefined)
        ? ['ods:auto-filter']
        : []),
      ...(document.workbook.sheets.some((sheet) => sheet.conditionalFormatting.length > 0)
        ? ['ods:conditional-formatting']
        : []),
      ...(document.workbook.sheets.some((sheet) => sheet.objects.length > 0)
        ? ['ods:drawing-objects']
        : []),
      ...validations.unsupported,
    ];
    return exportResult('ods', blob, unsupported);
  };
  return Object.freeze({
    format: 'ods',
    async write(document: SpreadsheetDocument, options = {}) {
      return (await writeResult(document, options as InterchangeWriteOptions)).blob;
    },
    writeResult(document: SpreadsheetDocument, options = {}) {
      return writeResult(document, options as InterchangeWriteOptions);
    },
  });
}
