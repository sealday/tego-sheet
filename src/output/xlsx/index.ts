import { strToU8, zipSync } from 'fflate';
import type { Cell, JsonValue, Sheet, SheetRange, SpreadsheetDocument } from '../../document';
import type { FormulaValue } from '../../formula';
import type { GeneratedDocument, GeneratedWorksheet, TemplatePrintProfile } from '../../template';
import { outputError, throwIfAborted } from '../output-error';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const FIXED_ZIP_DATE = new Date(1980, 0, 1, 0, 0, 0, 0);
const DEFAULT_LIMITS: XlsxOutputLimits = Object.freeze({
  maxSheets: 1_000,
  maxCells: 1_000_000,
  maxStyles: 65_000,
  maxImages: 10_000,
  maxStringBytes: 64 * 1024 * 1024,
  maxResourceBytes: 64 * 1024 * 1024,
  maxUncompressedBytes: 512 * 1024 * 1024,
  maxPackageBytes: 256 * 1024 * 1024,
});

/** XLSX conversion request. */
export interface XlsxOutputOptions {
  /** Formula preservation policy. */
  readonly formulaMode: 'formula-and-cached-value' | 'values-only';
  /** TP5 currently targets Excel-compatible OOXML. */
  readonly compatibility: 'excel';
  /** Cancels conversion without returning a partial package. */
  readonly signal?: AbortSignal;
}

/** Bounded semantic and package resources. */
export interface XlsxOutputLimits {
  /** Maximum worksheets. */
  readonly maxSheets: number;
  /** Maximum total sparse cells. */
  readonly maxCells: number;
  /** Maximum registered styles. */
  readonly maxStyles: number;
  /** Maximum generated worksheet images. */
  readonly maxImages: number;
  /** Maximum UTF-8 bytes across source strings. */
  readonly maxStringBytes: number;
  /** Maximum bytes across resolved resources. */
  readonly maxResourceBytes: number;
  /** Maximum total bytes across uncompressed package parts. */
  readonly maxUncompressedBytes: number;
  /** Maximum finalized ZIP bytes. */
  readonly maxPackageBytes: number;
}

/** XLSX adapter construction options. */
export interface XlsxAdapterOptions {
  /** Optional bounded output limits. */
  readonly limits?: Partial<XlsxOutputLimits>;
}

type Workbook = SpreadsheetDocument['workbook'];
type JsonRecord = Readonly<Record<string, JsonValue>>;

interface DrawingPart {
  readonly drawingPath: string;
  readonly relationshipsPath: string;
  readonly drawingXml: string;
  readonly relationshipsXml: string;
  readonly media: readonly {
    readonly path: string;
    readonly bytes: Uint8Array;
    readonly mimeType: 'image/png' | 'image/jpeg';
  }[];
}

function xml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function asRecord(value: JsonValue): JsonRecord | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}

function columnName(column: number): string {
  let value = column + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function cellReference(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`;
}

function absoluteRange(range: SheetRange): string {
  const start = `$${columnName(range.start.column)}$${range.start.row + 1}`;
  const end = `$${columnName(range.end.column)}$${range.end.row + 1}`;
  return `${start}:${end}`;
}

function sheetRange(sheet: Sheet): SheetRange {
  let endRow = Math.max(0, (sheet.rowCount ?? 1) - 1);
  let endColumn = Math.max(0, (sheet.columnCount ?? 1) - 1);
  for (const entry of sheet.cells) {
    endRow = Math.max(endRow, entry.row);
    endColumn = Math.max(endColumn, entry.column);
  }
  for (const merge of sheet.merges) {
    endRow = Math.max(endRow, merge.end.row);
    endColumn = Math.max(endColumn, merge.end.column);
  }
  return { start: { row: 0, column: 0 }, end: { row: endRow, column: endColumn } };
}

function quotedSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

function worksheetSettings(document: GeneratedDocument): ReadonlyMap<string, GeneratedWorksheet> {
  const settings = new Map<string, GeneratedWorksheet>();
  for (const worksheet of document.worksheets) {
    if (settings.has(worksheet.sheetId)) {
      throw outputError(
        'XLSX_UNSUPPORTED_FEATURE',
        `Worksheet settings for ${worksheet.sheetId} are duplicated`,
      );
    }
    settings.set(worksheet.sheetId, worksheet);
  }
  if (
    settings.size !== document.workbook.sheets.length ||
    document.workbook.sheets.some((sheet) => !settings.has(sheet.id)) ||
    !document.workbook.sheets.some((sheet) => settings.get(sheet.id)?.visibility === 'visible')
  ) {
    throw outputError(
      'XLSX_UNSUPPORTED_FEATURE',
      'XLSX requires exactly one setting per sheet and at least one visible sheet',
    );
  }
  return settings;
}

function xlsxColor(value: string): string {
  const normalized = value.startsWith('#') ? value.slice(1) : value;
  if (!/^(?:[\da-f]{6}|[\da-f]{8})$/iu.test(normalized)) {
    throw outputError('XLSX_UNSUPPORTED_FEATURE', `Conditional color ${value} is invalid`);
  }
  return normalized.length === 6 ? `FF${normalized.toUpperCase()}` : normalized.toUpperCase();
}

function conditionalFormattingXml(sheet: Sheet, worksheet: GeneratedWorksheet): string {
  return worksheet.conditionalFormatting
    .map((format, index) => {
      if (
        format.range.sheetId !== sheet.id ||
        format.range.start.row < 0 ||
        format.range.start.column < 0 ||
        format.range.end.row < format.range.start.row ||
        format.range.end.column < format.range.start.column
      ) {
        throw outputError('XLSX_UNSUPPORTED_FEATURE', 'Conditional formatting range is invalid', {
          location: { sheetId: sheet.id },
        });
      }
      const values =
        '<cfvo type="min"/>' +
        (format.midpointColor === undefined ? '' : '<cfvo type="percentile" val="50"/>') +
        '<cfvo type="max"/>';
      const colors =
        `<color rgb="${xlsxColor(format.minimumColor)}"/>` +
        (format.midpointColor === undefined
          ? ''
          : `<color rgb="${xlsxColor(format.midpointColor)}"/>`) +
        `<color rgb="${xlsxColor(format.maximumColor)}"/>`;
      return (
        `<conditionalFormatting sqref="${absoluteRange({ start: format.range.start, end: format.range.end }).replaceAll('$', '')}">` +
        `<cfRule type="colorScale" priority="${index + 1}"><colorScale>${values}${colors}</colorScale></cfRule>` +
        '</conditionalFormatting>'
      );
    })
    .join('');
}

function validateSheetNames(sheets: readonly Sheet[]): void {
  const names = new Set<string>();
  for (const sheet of sheets) {
    const normalized = sheet.name.toLocaleLowerCase('en-US');
    if (
      sheet.name.length === 0 ||
      sheet.name.length > 31 ||
      /[\\/?*:[\]]/u.test(sheet.name) ||
      names.has(normalized)
    ) {
      throw outputError(
        'XLSX_INVALID_SHEET_NAME',
        `Invalid or duplicate sheet name ${sheet.name}`,
        {
          location: { sheetId: sheet.id },
        },
      );
    }
    names.add(normalized);
  }
}

function calculatedValueMap(document: GeneratedDocument): ReadonlyMap<string, FormulaValue> {
  return new Map(
    document.calculatedCells.map(({ address, value }) => [
      `${address.sheetId}:${address.row}:${address.column}`,
      value,
    ]),
  );
}

function scalarValue(value: FormulaValue | undefined): string {
  if (value === undefined || value.type === 'blank') return '';
  if (value.type === 'array') {
    throw outputError(
      'XLSX_FORMULA_SERIALIZATION_FAILED',
      'Array formula results are not supported',
    );
  }
  if (value.type === 'boolean') return value.value ? '1' : '0';
  return xml(String(value.value));
}

function utf8Length(value: string): number {
  let bytes = 0;
  for (const character of value) {
    const point = character.codePointAt(0)!;
    bytes += point <= 0x7f ? 1 : point <= 0x7ff ? 2 : point <= 0xffff ? 3 : 4;
  }
  return bytes;
}

function jsonStringBytes(value: JsonValue): number {
  if (typeof value === 'string') return utf8Length(value);
  if (Array.isArray(value)) {
    return value.reduce((sum, entry) => sum + jsonStringBytes(entry), 0);
  }
  if (value !== null && typeof value === 'object') {
    return Object.entries(value).reduce(
      (sum, [key, entry]) => sum + utf8Length(key) + jsonStringBytes(entry),
      0,
    );
  }
  return 0;
}

function sourceStringBytes(document: GeneratedDocument): number {
  let total = 0;
  for (const sheet of document.workbook.sheets) {
    total += utf8Length(sheet.id) + utf8Length(sheet.name);
    for (const { cell } of sheet.cells) {
      const input = cell.input;
      if (input.type === 'string') total += utf8Length(input.value);
      if (input.type === 'formula') total += utf8Length(input.source);
      if (input.type === 'custom') {
        total += utf8Length(input.cellType) + jsonStringBytes(input.value);
      }
    }
  }
  for (const style of document.workbook.styles) {
    total += utf8Length(style.id) + jsonStringBytes(style.value);
  }
  for (const validation of document.workbook.validations) {
    total += utf8Length(validation.id) + jsonStringBytes(validation.value);
  }
  const { header, footer } = document.print.profile;
  for (const band of [header, footer]) {
    if (band === undefined) continue;
    for (const value of [band.left, band.center, band.right]) {
      if (value !== undefined) total += utf8Length(value);
    }
  }
  return total;
}

function conservativeUncompressedBytes(document: GeneratedDocument, stringBytes: number): number {
  const workbook = document.workbook;
  const cells = workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0);
  const rows = workbook.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const columns = workbook.sheets.reduce((sum, sheet) => sum + sheet.columns.length, 0);
  const merges = workbook.sheets.reduce((sum, sheet) => sum + sheet.merges.length, 0);
  const conditionalFormats = document.worksheets.reduce(
    (sum, worksheet) => sum + worksheet.conditionalFormatting.length,
    0,
  );
  return (
    64 * 1024 +
    stringBytes * 6 +
    document.resources.totalBytes +
    workbook.sheets.length * 8 * 1024 +
    cells * 256 +
    rows * 128 +
    columns * 128 +
    merges * 128 +
    workbook.styles.length * 2 * 1024 +
    workbook.validations.length * 1024 +
    document.objects.length * 2 * 1024 +
    conditionalFormats * 1024
  );
}

function cellXml(
  sheet: Sheet,
  row: number,
  column: number,
  cell: Cell,
  styleIndices: ReadonlyMap<string, number>,
  calculated: ReadonlyMap<string, FormulaValue>,
  formulaMode: XlsxOutputOptions['formulaMode'],
): string {
  const reference = cellReference(row, column);
  const style = cell.styleId === undefined ? undefined : styleIndices.get(cell.styleId);
  const styleAttribute = style === undefined ? '' : ` s="${style}"`;
  const input = cell.input;
  if (input.type === 'blank') return `<c r="${reference}"${styleAttribute}/>`;
  if (input.type === 'string') {
    return `<c r="${reference}" t="inlineStr"${styleAttribute}><is><t xml:space="preserve">${xml(input.value)}</t></is></c>`;
  }
  if (input.type === 'number') {
    return `<c r="${reference}"${styleAttribute}><v>${input.value}</v></c>`;
  }
  if (input.type === 'boolean') {
    return `<c r="${reference}" t="b"${styleAttribute}><v>${input.value ? 1 : 0}</v></c>`;
  }
  if (input.type === 'custom') {
    throw outputError(
      'XLSX_UNSUPPORTED_FEATURE',
      `Custom cell ${input.cellType} cannot be represented in XLSX`,
      {
        location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } },
        details: { cellType: input.cellType, schemaVersion: input.schemaVersion },
      },
    );
  }
  const cached = calculated.get(`${sheet.id}:${row}:${column}`);
  if (formulaMode === 'values-only') {
    if (cached === undefined || cached.type === 'blank') {
      return `<c r="${reference}"${styleAttribute}/>`;
    }
    if (cached.type === 'string') {
      return `<c r="${reference}" t="inlineStr"${styleAttribute}><is><t xml:space="preserve">${xml(cached.value)}</t></is></c>`;
    }
    if (cached.type === 'boolean') {
      return `<c r="${reference}" t="b"${styleAttribute}><v>${cached.value ? 1 : 0}</v></c>`;
    }
    if (cached.type === 'error') {
      return `<c r="${reference}" t="e"${styleAttribute}><v>${xml(cached.value)}</v></c>`;
    }
    if (cached.type === 'array') {
      throw outputError(
        'XLSX_FORMULA_SERIALIZATION_FAILED',
        `Array formula at ${reference} cannot be serialized`,
        { location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } } },
      );
    }
    return `<c r="${reference}"${styleAttribute}><v>${cached.value}</v></c>`;
  }
  const formula = input.source.startsWith('=') ? input.source.slice(1) : input.source;
  const cachedType =
    cached?.type === 'string'
      ? ' t="str"'
      : cached?.type === 'boolean'
        ? ' t="b"'
        : cached?.type === 'error'
          ? ' t="e"'
          : '';
  return `<c r="${reference}"${cachedType}${styleAttribute}><f>${xml(formula)}</f><v>${scalarValue(cached)}</v></c>`;
}

function styleXml(workbook: Workbook): {
  readonly xml: string;
  readonly indices: ReadonlyMap<string, number>;
} {
  const indices = new Map(workbook.styles.map((entry, index) => [entry.id, index + 1]));
  const numberFormats = workbook.styles
    .map((entry) => asRecord(entry.value)?.numberFormat)
    .filter((value): value is string => typeof value === 'string');
  const uniqueFormats = [...new Set(numberFormats)];
  const formatIds = new Map(uniqueFormats.map((format, index) => [format, 164 + index]));
  const fonts = workbook.styles.map((entry) => {
    const style = asRecord(entry.value);
    if (style === undefined) {
      throw outputError('XLSX_UNSUPPORTED_FEATURE', `Style ${entry.id} is not an object`);
    }
    const supported = new Set([
      'color',
      'backgroundColor',
      'fontFamily',
      'fontSize',
      'bold',
      'italic',
      'horizontalAlign',
      'verticalAlign',
      'wrap',
      'numberFormat',
      'underline',
      'strike',
      'border',
    ]);
    const unsupported = Object.keys(style).filter((key) => !supported.has(key));
    if (unsupported.length > 0) {
      throw outputError('XLSX_UNSUPPORTED_FEATURE', `Style ${entry.id} has unsupported fields`, {
        details: { styleId: entry.id, fields: unsupported },
      });
    }
    const name = typeof style.fontFamily === 'string' ? style.fontFamily : 'Arial';
    const size = typeof style.fontSize === 'number' ? style.fontSize : 10;
    const color = typeof style.color === 'string' ? style.color.replace('#', '') : 'FF000000';
    return `<font>${style.bold === true ? '<b/>' : ''}${style.italic === true ? '<i/>' : ''}${style.underline === true ? '<u/>' : ''}${style.strike === true ? '<strike/>' : ''}<sz val="${size}"/><color rgb="${color.length === 6 ? `FF${color}` : color}"/><name val="${xml(name)}"/></font>`;
  });
  const fills = workbook.styles.map((entry) => {
    const style = asRecord(entry.value)!;
    const color =
      typeof style.backgroundColor === 'string'
        ? style.backgroundColor.replace('#', '')
        : 'FFFFFFFF';
    return `<fill><patternFill patternType="solid"><fgColor rgb="${color.length === 6 ? `FF${color}` : color}"/><bgColor indexed="64"/></patternFill></fill>`;
  });
  const borders = workbook.styles.map((entry) => {
    const style = asRecord(entry.value)!;
    const border = style.border === undefined ? undefined : asRecord(style.border);
    const side = (name: string): string => {
      const value = border?.[name];
      if (!Array.isArray(value)) return `<${name}/>`;
      const lineStyle = typeof value[0] === 'string' ? value[0] : 'thin';
      const borderStyles = new Set([
        'hair',
        'dotted',
        'dashDotDot',
        'dashDot',
        'dashed',
        'thin',
        'mediumDashDotDot',
        'slantDashDot',
        'mediumDashDot',
        'mediumDashed',
        'medium',
        'thick',
        'double',
      ]);
      if (!borderStyles.has(lineStyle)) {
        throw outputError('XLSX_UNSUPPORTED_FEATURE', `Border style ${lineStyle} is unsupported`, {
          details: { styleId: entry.id, side: name, lineStyle },
        });
      }
      const color = typeof value[1] === 'string' ? value[1].replace('#', '') : 'FF000000';
      return `<${name} style="${xml(lineStyle)}"><color rgb="${color.length === 6 ? `FF${color}` : color}"/></${name}>`;
    };
    return `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`;
  });
  const xfs = workbook.styles.map((entry, index) => {
    const style = asRecord(entry.value)!;
    const format =
      typeof style.numberFormat === 'string' ? (formatIds.get(style.numberFormat) ?? 0) : 0;
    const horizontal =
      typeof style.horizontalAlign === 'string' ? style.horizontalAlign : 'general';
    const vertical = typeof style.verticalAlign === 'string' ? style.verticalAlign : 'bottom';
    if (!new Set(['general', 'left', 'center', 'right']).has(horizontal)) {
      throw outputError(
        'XLSX_UNSUPPORTED_FEATURE',
        `Horizontal alignment ${horizontal} is unsupported`,
        { details: { styleId: entry.id, horizontal } },
      );
    }
    if (!new Set(['top', 'middle', 'bottom']).has(vertical)) {
      throw outputError(
        'XLSX_UNSUPPORTED_FEATURE',
        `Vertical alignment ${vertical} is unsupported`,
        { details: { styleId: entry.id, vertical } },
      );
    }
    const xlsxVertical = vertical === 'middle' ? 'center' : vertical;
    return `<xf numFmtId="${format}" fontId="${index + 1}" fillId="${index + 2}" borderId="${index + 1}" xfId="0" applyFont="1" applyFill="1" applyBorder="1"${format === 0 ? '' : ' applyNumberFormat="1"'}><alignment horizontal="${horizontal}" vertical="${xlsxVertical}"${style.wrap === true ? ' wrapText="1"' : ''}/></xf>`;
  });
  return {
    indices,
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (uniqueFormats.length === 0
        ? ''
        : `<numFmts count="${uniqueFormats.length}">${uniqueFormats.map((format) => `<numFmt numFmtId="${formatIds.get(format)}" formatCode="${xml(format)}"/>`).join('')}</numFmts>`) +
      `<fonts count="${fonts.length + 1}"><font><sz val="10"/><name val="Arial"/></font>${fonts.join('')}</fonts>` +
      `<fills count="${fills.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills.join('')}</fills>` +
      `<borders count="${borders.length + 1}"><border><left/><right/><top/><bottom/><diagonal/></border>${borders.join('')}</borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
  };
}

function validationXml(sheet: Sheet, workbook: Workbook): string {
  const definitions = new Map(workbook.validations.map((entry) => [entry.id, entry.value]));
  const validations = sheet.cells.flatMap(({ row, column, cell }) => {
    if (cell.validationId === undefined) return [];
    const definition = definitions.get(cell.validationId);
    const rule = definition === undefined ? undefined : asRecord(definition);
    if (rule === undefined || typeof rule.type !== 'string') {
      throw outputError('XLSX_UNSUPPORTED_FEATURE', 'Cell validation is unsupported', {
        location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } },
      });
    }
    const supportedTypes = new Set([
      'whole',
      'decimal',
      'list',
      'date',
      'time',
      'textLength',
      'custom',
    ]);
    if (!supportedTypes.has(rule.type)) {
      throw outputError('XLSX_UNSUPPORTED_FEATURE', `Validation type ${rule.type} is unsupported`, {
        location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } },
      });
    }
    const operator = typeof rule.operator === 'string' ? ` operator="${xml(rule.operator)}"` : '';
    if (
      typeof rule.operator === 'string' &&
      !new Set([
        'between',
        'notBetween',
        'equal',
        'notEqual',
        'greaterThan',
        'lessThan',
        'greaterThanOrEqual',
        'lessThanOrEqual',
      ]).has(rule.operator)
    ) {
      throw outputError(
        'XLSX_UNSUPPORTED_FEATURE',
        `Validation operator ${rule.operator} is unsupported`,
        { location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } } },
      );
    }
    const allowBlank = rule.allowBlank === true ? ' allowBlank="1"' : '';
    const formula1 =
      typeof rule.formula1 === 'string' || typeof rule.formula1 === 'number'
        ? `<formula1>${xml(String(rule.formula1))}</formula1>`
        : '';
    const formula2 =
      typeof rule.formula2 === 'string' || typeof rule.formula2 === 'number'
        ? `<formula2>${xml(String(rule.formula2))}</formula2>`
        : '';
    return [
      `<dataValidation type="${xml(rule.type)}"${operator}${allowBlank} sqref="${cellReference(row, column)}">${formula1}${formula2}</dataValidation>`,
    ];
  });
  return validations.length === 0
    ? ''
    : `<dataValidations count="${validations.length}">${validations.join('')}</dataValidations>`;
}

function pageXml(profile: TemplatePrintProfile | undefined, sheet: Sheet): string {
  if (profile === undefined) return '';
  const margins = profile.page.margins;
  const paper =
    profile.page.paper.type === 'A4'
      ? 9
      : profile.page.paper.type === 'A5'
        ? 11
        : profile.page.paper.type === 'Letter'
          ? 1
          : undefined;
  const paperGeometry =
    profile.page.paper.type === 'custom'
      ? ` paperWidth="${profile.page.paper.width / 96}in" paperHeight="${profile.page.paper.height / 96}in"`
      : '';
  const scale = profile.page.scale;
  const setup =
    scale.type === 'fixed'
      ? ` scale="${Math.max(10, Math.min(400, Math.round(scale.value * 100)))}"`
      : ` fitToWidth="${scale.type === 'fit-width' ? scale.pages : 1}" fitToHeight="${scale.type === 'fit-page' ? 1 : 0}"`;
  const breaks = profile.manualBreaks.filter(({ sheetId }) => sheetId === sheet.id);
  const header = profile.header;
  const footer = profile.footer;
  const band = (value: typeof header): string =>
    value === undefined
      ? ''
      : `${value.left === undefined ? '' : `&L${value.left}`}${value.center === undefined ? '' : `&C${value.center}`}${value.right === undefined ? '' : `&R${value.right}`}`;
  return (
    `<pageMargins left="${margins.left / 96}" right="${margins.right / 96}" top="${margins.top / 96}" bottom="${margins.bottom / 96}" header="0" footer="0"/>` +
    `<pageSetup orientation="${profile.page.orientation}"${paper === undefined ? '' : ` paperSize="${paper}"`}${paperGeometry}${setup}/>` +
    (breaks.length === 0
      ? ''
      : `<rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">${breaks.map(({ beforeRow }) => `<brk id="${beforeRow}" min="0" max="16383" man="1"/>`).join('')}</rowBreaks>`) +
    (header === undefined && footer === undefined
      ? ''
      : `<headerFooter><oddHeader>${xml(band(header))}</oddHeader><oddFooter>${xml(band(footer))}</oddFooter></headerFooter>`)
  );
}

function drawingPart(
  sheet: Sheet,
  sheetIndex: number,
  document: GeneratedDocument,
): DrawingPart | undefined {
  const objects = document.objects.filter(({ generated }) => generated.sheetId === sheet.id);
  if (objects.length === 0) return undefined;
  const mediaByHash = new Map<
    string,
    {
      readonly path: string;
      readonly bytes: Uint8Array;
      readonly mimeType: 'image/png' | 'image/jpeg';
    }
  >();
  const relationships: string[] = [];
  const anchors = objects.map((object, index) => {
    if (object.resourceId === undefined) {
      throw outputError('XLSX_RESOURCE_UNSUPPORTED', `Object ${object.objectId} has no resource`, {
        details: { objectId: object.objectId },
      });
    }
    const resource = document.resources.byReference[object.resourceId];
    if (
      resource === undefined ||
      (resource.mimeType !== 'image/png' && resource.mimeType !== 'image/jpeg')
    ) {
      throw outputError(
        'XLSX_RESOURCE_UNSUPPORTED',
        `Object ${object.objectId} is not PNG or JPEG`,
        {
          details: { objectId: object.objectId, resourceId: object.resourceId },
        },
      );
    }
    let media = mediaByHash.get(resource.contentHash);
    if (media === undefined) {
      const extension = resource.mimeType === 'image/png' ? 'png' : 'jpeg';
      media = {
        path: `xl/media/image${sheetIndex + 1}-${mediaByHash.size + 1}.${extension}`,
        bytes: new Uint8Array(resource.bytes),
        mimeType: resource.mimeType,
      };
      mediaByHash.set(resource.contentHash, media);
    }
    const relationshipId = `rId${index + 1}`;
    relationships.push(
      `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${media.path.split('/').at(-1)}"/>`,
    );
    const range = object.generated;
    return (
      `<xdr:twoCellAnchor><xdr:from><xdr:col>${range.start.column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${range.start.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>${range.end.column + 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${range.end.row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 1}" name="${xml(object.objectId)}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`
    );
  });
  return {
    drawingPath: `xl/drawings/drawing${sheetIndex + 1}.xml`,
    relationshipsPath: `xl/drawings/_rels/drawing${sheetIndex + 1}.xml.rels`,
    drawingXml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `${anchors.join('')}</xdr:wsDr>`,
    relationshipsXml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`,
    media: [...mediaByHash.values()],
  };
}

function worksheetXml(
  sheet: Sheet,
  workbook: Workbook,
  styleIndices: ReadonlyMap<string, number>,
  calculated: ReadonlyMap<string, FormulaValue>,
  options: XlsxOutputOptions,
  profile: TemplatePrintProfile | undefined,
  worksheet: GeneratedWorksheet,
  hasDrawing: boolean,
): string {
  const rows = new Map<
    number,
    { readonly cells: typeof sheet.cells; readonly height?: number; readonly hidden?: boolean }
  >();
  for (const row of sheet.rows) {
    rows.set(row.index, {
      cells: [],
      ...(row.height === undefined ? {} : { height: row.height }),
      ...(row.hidden === undefined ? {} : { hidden: row.hidden }),
    });
  }
  for (const entry of sheet.cells) {
    const current = rows.get(entry.row) ?? { cells: [] };
    rows.set(entry.row, { ...current, cells: [...current.cells, entry] });
  }
  const columns =
    sheet.columns.length === 0
      ? ''
      : `<cols>${sheet.columns
          .map(
            (column) =>
              `<col min="${column.index + 1}" max="${column.index + 1}"${column.width === undefined ? '' : ` width="${column.width}" customWidth="1"`}${column.hidden === true ? ' hidden="1"' : ''}/>`,
          )
          .join('')}</cols>`;
  const rowXml = [...rows.entries()]
    .sort(([left], [right]) => left - right)
    .map(
      ([row, value]) =>
        `<row r="${row + 1}"${value.height === undefined ? '' : ` ht="${value.height}" customHeight="1"`}${value.hidden === true ? ' hidden="1"' : ''}>${value.cells
          .map((entry) =>
            cellXml(
              sheet,
              entry.row,
              entry.column,
              entry.cell,
              styleIndices,
              calculated,
              options.formulaMode,
            ),
          )
          .join('')}</row>`,
    )
    .join('');
  const merges =
    sheet.merges.length === 0
      ? ''
      : `<mergeCells count="${sheet.merges.length}">${sheet.merges
          .map(
            (range) =>
              `<mergeCell ref="${cellReference(range.start.row, range.start.column)}:${cellReference(range.end.row, range.end.column)}"/>`,
          )
          .join('')}</mergeCells>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="${absoluteRange(sheetRange(sheet)).replaceAll('$', '')}"/>` +
    `<sheetViews><sheetView workbookViewId="0"${profile?.showGridlines === false ? ' showGridLines="0"' : ''}${profile?.showHeadings === false ? ' showRowColHeaders="0"' : ''}/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/>${columns}<sheetData>${rowXml}</sheetData>${merges}` +
    `${validationXml(sheet, workbook)}${conditionalFormattingXml(sheet, worksheet)}${pageXml(profile, sheet)}${hasDrawing ? '<drawing r:id="rId1"/>' : ''}</worksheet>`
  );
}

function workbookDefinedNames(
  workbook: Workbook,
  profile: TemplatePrintProfile | undefined,
): string {
  if (profile === undefined) return '';
  const names = workbook.sheets.flatMap((sheet, index) => {
    const result: string[] = [];
    const printAreas = profile.targets.flatMap((target): readonly SheetRange[] => {
      if (target.type === 'sheet') return target.sheetId === sheet.id ? [sheetRange(sheet)] : [];
      if (target.type === 'range') {
        return target.range.sheetId === sheet.id
          ? [{ start: target.range.start, end: target.range.end }]
          : [];
      }
      return target.ranges
        .filter((range) => range.sheetId === sheet.id)
        .map((range) => ({ start: range.start, end: range.end }));
    });
    if (printAreas.length > 0) {
      result.push(
        `<definedName name="_xlnm.Print_Area" localSheetId="${index}">${printAreas.map((range) => `${quotedSheetName(sheet.name)}!${absoluteRange(range)}`).join(',')}</definedName>`,
      );
    }
    const titleRows = profile.repeatRows?.sheetId === sheet.id ? profile.repeatRows : undefined;
    const titleColumns =
      profile.repeatColumns?.sheetId === sheet.id ? profile.repeatColumns : undefined;
    if (titleRows !== undefined || titleColumns !== undefined) {
      const rows =
        titleRows === undefined
          ? ''
          : `${quotedSheetName(sheet.name)}!$${titleRows.start.row + 1}:$${titleRows.end.row + 1}`;
      const columns =
        titleColumns === undefined
          ? ''
          : `${quotedSheetName(sheet.name)}!$${columnName(titleColumns.start.column)}:$${columnName(titleColumns.end.column)}`;
      result.push(
        `<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${rows}${rows !== '' && columns !== '' ? ',' : ''}${columns}</definedName>`,
      );
    }
    return result;
  });
  return names.length === 0 ? '' : `<definedNames>${names.join('')}</definedNames>`;
}

function contentTypes(drawings: readonly (DrawingPart | undefined)[]): string {
  const mediaTypes = new Set(
    drawings.flatMap((drawing) => drawing?.media.map(({ mimeType }) => mimeType) ?? []),
  );
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `${mediaTypes.has('image/png') ? '<Default Extension="png" ContentType="image/png"/>' : ''}` +
    `${mediaTypes.has('image/jpeg') ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : ''}` +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    `${drawings.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}` +
    `${drawings.flatMap((drawing) => (drawing === undefined ? [] : [`<Override PartName="/${drawing.drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`])).join('')}` +
    '</Types>'
  );
}

function packageParts(
  document: GeneratedDocument,
  options: XlsxOutputOptions,
): Map<string, Uint8Array> {
  const workbook = document.workbook;
  validateSheetNames(workbook.sheets);
  for (const sheet of workbook.sheets) {
    for (const { row, column, cell } of sheet.cells) {
      const metadata = cell.metadata === undefined ? undefined : asRecord(cell.metadata);
      const conditional =
        metadata?.conditionalFormatting ?? metadata?.['tego-sheet:conditionalFormatting'];
      if (conditional !== undefined) {
        throw outputError(
          'XLSX_UNSUPPORTED_FEATURE',
          'Conditional formatting metadata is not yet a Workbook 2.0 contract',
          { location: { sheetId: sheet.id, cell: { sheetId: sheet.id, row, column } } },
        );
      }
    }
  }
  const style = styleXml(workbook);
  const calculated = calculatedValueMap(document);
  const profile = document.print.profile;
  const worksheetById = worksheetSettings(document);
  const drawings = workbook.sheets.map((sheet, index) => drawingPart(sheet, index, document));
  const parts = new Map<string, Uint8Array>();
  const add = (path: string, value: string | Uint8Array): void => {
    parts.set(path, typeof value === 'string' ? strToU8(value) : value);
  };
  add('[Content_Types].xml', contentTypes(drawings));
  add(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  );
  add(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<workbookPr date1904="${workbook.settings.dateSystem === 'excel-1904' ? 1 : 0}"/><bookViews><workbookView/></bookViews><sheets>${workbook.sheets
        .map((sheet, index) => {
          const visibility = worksheetById.get(sheet.id)!.visibility;
          const state =
            visibility === 'visible'
              ? ''
              : ` state="${visibility === 'very-hidden' ? 'veryHidden' : 'hidden'}"`;
          return `<sheet name="${xml(sheet.name)}" sheetId="${index + 1}"${state} r:id="rId${index + 1}"/>`;
        })
        .join(
          '',
        )}</sheets>${workbookDefinedNames(workbook, profile)}<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`,
  );
  add(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbook.sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('')}<Relationship Id="rId${workbook.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  add('xl/styles.xml', style.xml);
  for (const [index, sheet] of workbook.sheets.entries()) {
    throwIfAborted(options.signal);
    const drawing = drawings[index];
    add(
      `xl/worksheets/sheet${index + 1}.xml`,
      worksheetXml(
        sheet,
        workbook,
        style.indices,
        calculated,
        options,
        profile,
        worksheetById.get(sheet.id)!,
        drawing !== undefined,
      ),
    );
    if (drawing !== undefined) {
      add(drawing.drawingPath, drawing.drawingXml);
      add(drawing.relationshipsPath, drawing.relationshipsXml);
      add(
        `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/' +
          `drawing${index + 1}.xml"/></Relationships>`,
      );
      for (const media of drawing.media) add(media.path, media.bytes);
    }
  }
  return parts;
}

/** Pure write-only OOXML translation of an immutable generated document. */
export class XlsxAdapter {
  readonly #limits: XlsxOutputLimits;

  /** Creates a bounded XLSX adapter. */
  constructor(options: XlsxAdapterOptions = {}) {
    this.#limits = Object.freeze({ ...DEFAULT_LIMITS, ...options.limits });
  }

  /** Renders one atomic deterministic XLSX package. */
  async render(document: GeneratedDocument, options: XlsxOutputOptions): Promise<Blob> {
    throwIfAborted(options.signal);
    const workbook = document.workbook;
    const cells = workbook.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0);
    const stringBytes = sourceStringBytes(document);
    const estimatedUncompressedBytes = conservativeUncompressedBytes(document, stringBytes);
    if (
      workbook.sheets.length > this.#limits.maxSheets ||
      cells > this.#limits.maxCells ||
      workbook.styles.length > this.#limits.maxStyles ||
      document.objects.length > this.#limits.maxImages ||
      stringBytes > this.#limits.maxStringBytes ||
      document.resources.totalBytes > this.#limits.maxResourceBytes ||
      !Number.isSafeInteger(estimatedUncompressedBytes) ||
      estimatedUncompressedBytes > this.#limits.maxUncompressedBytes
    ) {
      throw outputError('XLSX_PACKAGE_LIMIT_EXCEEDED', 'XLSX semantic limits were exceeded');
    }
    const parts = packageParts(document, options);
    const uncompressedBytes = [...parts.values()].reduce(
      (total, bytes) => total + bytes.byteLength,
      0,
    );
    if (
      !Number.isSafeInteger(uncompressedBytes) ||
      uncompressedBytes > this.#limits.maxUncompressedBytes
    ) {
      throw outputError(
        'XLSX_PACKAGE_LIMIT_EXCEEDED',
        'XLSX uncompressed parts exceed their byte limit',
      );
    }
    throwIfAborted(options.signal);
    const ordered: Parameters<typeof zipSync>[0] = {};
    for (const [path, bytes] of [...parts.entries()].sort(([left], [right]) =>
      left < right ? -1 : left > right ? 1 : 0,
    )) {
      ordered[path] = [bytes, { level: 6, mtime: FIXED_ZIP_DATE }];
    }
    const bytes = zipSync(ordered, { level: 6, mtime: FIXED_ZIP_DATE });
    throwIfAborted(options.signal);
    if (bytes.byteLength > this.#limits.maxPackageBytes) {
      throw outputError('XLSX_PACKAGE_LIMIT_EXCEEDED', 'XLSX package exceeds its byte limit');
    }
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    return new Blob([buffer], { type: XLSX_MIME });
  }
}

export { OutputAdapterError } from '../output-error';
