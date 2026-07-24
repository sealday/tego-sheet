import type {
  CellInput,
  ConditionalStyle,
  DocumentSheetId,
  JsonValue,
  ResourceMetadata,
  SpreadsheetDocumentInput,
} from '../document';
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
import { createXlsxDrawingResourcePool, parseWorksheetDrawing } from './xlsx-drawing';

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

function sheetRange(reference: string): {
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
} {
  const [startReference, endReference = startReference] = reference.replaceAll('$', '').split(':');
  if (startReference === undefined || endReference === undefined) {
    throw new InterchangeError('MALFORMED_WORKBOOK', `Invalid range reference: ${reference}`);
  }
  return {
    start: { row: rowIndex(startReference), column: columnIndex(startReference) },
    end: { row: rowIndex(endReference), column: columnIndex(endReference) },
  };
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

interface ParsedXlsxStyles {
  readonly registry: readonly { readonly id: string; readonly value: JsonValue }[];
  readonly cellStyleIds: ReadonlyMap<number, string>;
  readonly differentialStyles: readonly ConditionalStyle[];
}

function children(xml: string, element: string): readonly string[] {
  const container = new RegExp(`<${element}s\\b[^>]*>([\\s\\S]*?)<\\/${element}s>`, 'i').exec(
    xml,
  )?.[1];
  if (container === undefined) return [];
  return [
    ...container.matchAll(
      new RegExp(`<${element}\\b[^>]*?(?:\\/>|>([\\s\\S]*?)<\\/${element}>)`, 'gi'),
    ),
  ].map((match) => match[0]);
}

function normalizedColor(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const normalized = value.startsWith('FF') && value.length === 8 ? value.slice(2) : value;
  return /^[\dA-F]{6}$/i.test(normalized) ? `#${normalized.toLowerCase()}` : undefined;
}

function styleValue(
  xfXml: string,
  fonts: readonly string[],
  fills: readonly string[],
  borders: readonly string[],
  numberFormats: ReadonlyMap<number, string>,
): JsonValue {
  const xf = attributes(/^<xf\b([^>]*)/i.exec(xfXml)?.[1] ?? '');
  const font = fonts[Number(xf.fontId ?? 0)] ?? '';
  const fill = fills[Number(xf.fillId ?? 0)] ?? '';
  const border = borders[Number(xf.borderId ?? 0)] ?? '';
  const alignmentXml = /<alignment\b([^>]*?)(?:\/>|>)/i.exec(xfXml)?.[1];
  const alignment = alignmentXml === undefined ? {} : attributes(alignmentXml);
  const result: Record<string, JsonValue> = {};
  const fontColor = normalizedColor(
    attributes(/<color\b([^>]*?)(?:\/>|>)/i.exec(font)?.[1] ?? '').rgb,
  );
  const fillColor = normalizedColor(
    attributes(/<fgColor\b([^>]*?)(?:\/>|>)/i.exec(fill)?.[1] ?? '').rgb,
  );
  const name = attributes(/<name\b([^>]*?)(?:\/>|>)/i.exec(font)?.[1] ?? '').val;
  const size = Number(attributes(/<sz\b([^>]*?)(?:\/>|>)/i.exec(font)?.[1] ?? '').val);
  if (fontColor !== undefined) result.color = fontColor;
  if (fillColor !== undefined) result.backgroundColor = fillColor;
  if (name !== undefined) result.fontFamily = name;
  if (Number.isFinite(size)) result.fontSize = size;
  if (/<b(?:\s[^>]*)?\/?>/i.test(font)) result.bold = true;
  if (/<i(?:\s[^>]*)?\/?>/i.test(font)) result.italic = true;
  if (alignment.horizontal !== undefined) result.horizontalAlign = alignment.horizontal;
  if (alignment.vertical !== undefined) {
    result.verticalAlign = alignment.vertical === 'center' ? 'middle' : alignment.vertical;
  }
  if (alignment.wrapText === '1') result.wrap = true;
  const numberFormat = numberFormats.get(Number(xf.numFmtId ?? 0));
  if (numberFormat !== undefined) result.numberFormat = numberFormat;
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const match = new RegExp(`<${side}\\b([^>]*)>([\\s\\S]*?)<\\/${side}>`, 'i').exec(border);
    if (match === null) continue;
    const lineStyle = attributes(match[1]!).style;
    const color = normalizedColor(
      attributes(/<color\b([^>]*?)(?:\/>|>)/i.exec(match[2]!)?.[1] ?? '').rgb,
    );
    if (lineStyle !== undefined && color !== undefined)
      result[`${side}Border`] = [lineStyle, color];
  }
  return result;
}

function parseStyles(xml: string | undefined): ParsedXlsxStyles {
  if (xml === undefined) {
    return { registry: [], cellStyleIds: new Map(), differentialStyles: [] };
  }
  const numberFormats = new Map(
    [...xml.matchAll(/<numFmt\b([^>]*?)(?:\/>|>[\s\S]*?<\/numFmt>)/gi)].flatMap((match) => {
      const entry = attributes(match[1]!);
      const id = Number(entry.numFmtId);
      return Number.isSafeInteger(id) && entry.formatCode !== undefined
        ? ([[id, entry.formatCode]] as const)
        : [];
    }),
  );
  const fonts = children(xml, 'font');
  const fills = children(xml, 'fill');
  const borders = children(xml, 'border');
  const cellXfsBody = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i.exec(xml)?.[1] ?? '';
  const xfs = [...cellXfsBody.matchAll(/<xf\b[^>]*?(?:\/>|>([\s\S]*?)<\/xf>)/gi)].map(
    (match) => match[0],
  );
  const registry = xfs.slice(1).map((xf, index) => ({
    id: `xlsx-style-${index + 1}`,
    value: styleValue(xf, fonts, fills, borders, numberFormats),
  }));
  const cellStyleIds = new Map(registry.map((entry, index) => [index + 1, entry.id]));
  const differentialStyles = children(xml, 'dxf').map((dxf) => {
    const font = /<font\b[^>]*>([\s\S]*?)<\/font>/i.exec(dxf)?.[1] ?? '';
    const fill = /<fill\b[^>]*>([\s\S]*?)<\/fill>/i.exec(dxf)?.[1] ?? '';
    const color = normalizedColor(
      attributes(/<color\b([^>]*?)(?:\/>|>)/i.exec(font)?.[1] ?? '').rgb,
    );
    const backgroundColor = normalizedColor(
      attributes(/<fgColor\b([^>]*?)(?:\/>|>)/i.exec(fill)?.[1] ?? '').rgb,
    );
    return {
      ...(color === undefined ? {} : { color }),
      ...(backgroundColor === undefined ? {} : { backgroundColor }),
      ...(/<b(?:\s[^>]*)?\/?>/i.test(font) ? { bold: true } : {}),
    };
  });
  return { registry, cellStyleIds, differentialStyles };
}

function pageBand(value: string): { left?: string; center?: string; right?: string } | undefined {
  if (value === '') return undefined;
  const result: { left?: string; center?: string; right?: string } = {};
  const matches = [...value.matchAll(/&([LCR])/g)];
  for (const [index, match] of matches.entries()) {
    const content = value.slice(match.index + 2, matches[index + 1]?.index ?? value.length);
    if (match[1] === 'L') result.left = content;
    if (match[1] === 'C') result.center = content;
    if (match[1] === 'R') result.right = content;
  }
  return Object.keys(result).length === 0 ? { center: value } : result;
}

function printProfile(
  sheetId: string,
  xml: string,
): SpreadsheetDocumentInput['templates'][number]['printProfiles'][number] | undefined {
  if (!/<(?:pageMargins|pageSetup|headerFooter|printOptions)\b/i.test(xml)) return undefined;
  const marginAttributes = attributes(/<pageMargins\b([^>]*?)(?:\/>|>)/i.exec(xml)?.[1] ?? '');
  const setup = attributes(/<pageSetup\b([^>]*?)(?:\/>|>)/i.exec(xml)?.[1] ?? '');
  const view = attributes(/<sheetView\b([^>]*?)(?:\/>|>)/i.exec(xml)?.[1] ?? '');
  const paper =
    setup.paperSize === '11'
      ? ({ type: 'A5' } as const)
      : setup.paperSize === '1'
        ? ({ type: 'Letter' } as const)
        : setup.paperWidth !== undefined && setup.paperHeight !== undefined
          ? ({
              type: 'custom',
              width: Number.parseFloat(setup.paperWidth) * 96,
              height: Number.parseFloat(setup.paperHeight) * 96,
            } as const)
          : ({ type: 'A4' } as const);
  const scale =
    setup.scale !== undefined
      ? ({ type: 'fixed', value: Number(setup.scale) / 100 } as const)
      : setup.fitToHeight === '1'
        ? ({ type: 'fit-page' } as const)
        : ({ type: 'fit-width', pages: Number(setup.fitToWidth ?? 1) } as const);
  const headerFooter = /<headerFooter\b[^>]*>([\s\S]*?)<\/headerFooter>/i.exec(xml)?.[1] ?? '';
  const header = pageBand(
    textContent(/<oddHeader\b[^>]*>([\s\S]*?)<\/oddHeader>/i.exec(headerFooter)?.[1] ?? ''),
  );
  const footer = pageBand(
    textContent(/<oddFooter\b[^>]*>([\s\S]*?)<\/oddFooter>/i.exec(headerFooter)?.[1] ?? ''),
  );
  const documentSheetId = sheetId as DocumentSheetId;
  return {
    id: `xlsx-print-${sheetId}`,
    name: 'Imported print',
    targets: [{ type: 'sheet', sheetId: documentSheetId }],
    page: {
      paper,
      orientation: setup.orientation === 'landscape' ? 'landscape' : 'portrait',
      margins: {
        top: Number(marginAttributes.top ?? 0) * 96,
        right: Number(marginAttributes.right ?? 0) * 96,
        bottom: Number(marginAttributes.bottom ?? 0) * 96,
        left: Number(marginAttributes.left ?? 0) * 96,
      },
      scale,
    },
    manualBreaks: [...xml.matchAll(/<brk\b([^>]*?)(?:\/>|>)/gi)].flatMap((match) => {
      const beforeRow = Number(attributes(match[1]!).id);
      return Number.isSafeInteger(beforeRow) ? [{ sheetId: documentSheetId, beforeRow }] : [];
    }),
    ...(header === undefined ? {} : { header }),
    ...(footer === undefined ? {} : { footer }),
    showGridlines: view.showGridLines !== '0',
    showHeadings: view.showRowColHeaders !== '0',
  };
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
  id: string,
  name: string,
  xml: string,
  strings: readonly string[],
  styles: ParsedXlsxStyles,
  validationRegistry: { id: string; value: JsonValue }[],
  limits: ReturnType<typeof resolveLimits>,
  unsupported: string[],
  visibility: ImportedSheet['visibility'],
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
      const styleId =
        cellAttributes.s === undefined
          ? undefined
          : styles.cellStyleIds.get(Number(cellAttributes.s));
      cells.push({ row, column, input, ...(styleId === undefined ? {} : { styleId }) });
      if (cells.length > limits.maxCells) {
        throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'XLSX cell limit exceeded');
      }
    }
  }
  const merges = [...xml.matchAll(/<mergeCell\b([^>]*?)(?:\/>|>[\s\S]*?<\/mergeCell>)/gi)].map(
    (match) => {
      const reference = attributes(match[1]!).ref;
      if (reference === undefined) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'Merged range reference is missing');
      }
      return sheetRange(reference);
    },
  );
  const byCoordinate = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  const setValidation = (reference: string, validationId: string): void => {
    const range = sheetRange(reference);
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        const key = `${row}:${column}`;
        const current = byCoordinate.get(key);
        const next = {
          row,
          column,
          input: current?.input ?? ({ type: 'blank' } as const),
          ...(current?.styleId === undefined ? {} : { styleId: current.styleId }),
          validationId,
        };
        if (current === undefined) cells.push(next);
        else cells[cells.indexOf(current)] = next;
        byCoordinate.set(key, next);
      }
    }
  };
  for (const match of xml.matchAll(
    /<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/gi,
  )) {
    const ruleAttributes = attributes(match[1]!);
    if (ruleAttributes.sqref === undefined || ruleAttributes.type === undefined) {
      throw new InterchangeError('MALFORMED_WORKBOOK', 'Data validation is incomplete');
    }
    const body = match[2] ?? '';
    const scalar = (tag: 'formula1' | 'formula2'): JsonValue | undefined => {
      const value = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)<\\/${tag}>`, 'i').exec(body)?.[1];
      if (value === undefined) return undefined;
      const decoded = textContent(value);
      const number = Number(decoded);
      return Number.isFinite(number) && decoded.trim() !== '' ? number : decoded;
    };
    const formula1 = scalar('formula1');
    const formula2 = scalar('formula2');
    const id = `xlsx-validation-${validationRegistry.length + 1}`;
    validationRegistry.push({
      id,
      value: {
        type: ruleAttributes.type,
        ...(ruleAttributes.operator === undefined ? {} : { operator: ruleAttributes.operator }),
        ...(formula1 === undefined ? {} : { formula1 }),
        ...(formula2 === undefined ? {} : { formula2 }),
        allowBlank: ruleAttributes.allowBlank === '1',
      },
    });
    for (const reference of ruleAttributes.sqref.split(/\s+/).filter(Boolean)) {
      setValidation(reference, id);
    }
  }
  const conditionalFormatting: NonNullable<ImportedSheet['conditionalFormatting']>[number][] = [];
  for (const match of xml.matchAll(
    /<conditionalFormatting\b([^>]*?)>([\s\S]*?)<\/conditionalFormatting>/gi,
  )) {
    const reference = attributes(match[1]!).sqref;
    if (reference === undefined || /\s/.test(reference.trim())) {
      unsupported.push('xlsx:conditional-formatting-multi-range');
      continue;
    }
    const range = { sheetId: id as DocumentSheetId, ...sheetRange(reference) };
    const rule = /<cfRule\b([^>]*?)>([\s\S]*?)<\/cfRule>/i.exec(match[2]!);
    if (rule === null) continue;
    const ruleAttributes = attributes(rule[1]!);
    if (ruleAttributes.type === 'colorScale') {
      const colors = [...rule[2]!.matchAll(/<color\b([^>]*?)(?:\/>|>)/gi)]
        .map((color) => normalizedColor(attributes(color[1]!).rgb))
        .filter((color): color is string => color !== undefined);
      if (colors.length === 2 || colors.length === 3) {
        conditionalFormatting.push({
          type: 'color-scale',
          range,
          minimumColor: colors[0]!,
          ...(colors.length === 3 ? { midpointColor: colors[1]! } : {}),
          maximumColor: colors.at(-1)!,
        });
      } else unsupported.push('xlsx:conditional-formatting-color');
    } else if (ruleAttributes.type === 'cellIs' && ruleAttributes.operator !== undefined) {
      const formulas = [...rule[2]!.matchAll(/<formula\b[^>]*>([\s\S]*?)<\/formula>/gi)].map(
        (formula) => textContent(formula[1]!),
      );
      const differential =
        ruleAttributes.dxfId === undefined
          ? {}
          : (styles.differentialStyles[Number(ruleAttributes.dxfId)] ?? {});
      conditionalFormatting.push({
        type: 'cell-is',
        range,
        operator: ruleAttributes.operator as never,
        formula: formulas[0] ?? '',
        ...(formulas[1] === undefined ? {} : { formula2: formulas[1] }),
        style: differential,
      });
    } else {
      unsupported.push(`xlsx:conditional-formatting-${ruleAttributes.type ?? 'unknown'}`);
    }
  }
  const autoFilter = /<autoFilter\b([^>]*?)(?:\/>|>([\s\S]*?)<\/autoFilter>)/i.exec(xml);
  let filter: ImportedSheet['filter'];
  if (autoFilter !== null) {
    const reference = attributes(autoFilter[1]!).ref;
    if (reference === undefined) {
      throw new InterchangeError('MALFORMED_WORKBOOK', 'Auto-filter range is missing');
    }
    const range = sheetRange(reference);
    const body = autoFilter[2] ?? '';
    const filters = [...body.matchAll(/<filterColumn\b([^>]*?)>([\s\S]*?)<\/filterColumn>/gi)].map(
      (match) => {
        const relativeColumn = Number(attributes(match[1]!).colId);
        if (!Number.isSafeInteger(relativeColumn) || relativeColumn < 0) {
          throw new InterchangeError('MALFORMED_WORKBOOK', 'Auto-filter column is invalid');
        }
        return {
          column: range.start.column + relativeColumn,
          operator: 'in' as const,
          values: [...match[2]!.matchAll(/<filter\b([^>]*?)(?:\/>|>[\s\S]*?<\/filter>)/gi)].map(
            (value) => attributes(value[1]!).val ?? '',
          ),
        };
      },
    );
    const sortMatch = /<sortCondition\b([^>]*?)(?:\/>|>[\s\S]*?<\/sortCondition>)/i.exec(body);
    const sortAttributes = sortMatch === null ? undefined : attributes(sortMatch[1]!);
    const sortReference = sortAttributes?.ref;
    filter = {
      range,
      filters,
      ...(sortReference === undefined
        ? {}
        : {
            sort: {
              column: sheetRange(sortReference).start.column,
              direction: sortAttributes?.descending === '1' ? ('desc' as const) : ('asc' as const),
            },
          }),
    };
  }
  if (/<hyperlinks\b/i.test(xml)) unsupported.push('xlsx:hyperlinks');
  if (/<sheetProtection\b/i.test(xml)) unsupported.push('xlsx:sheet-protection');
  if (/<drawing\b/i.test(xml)) unsupported.push('xlsx:drawing-objects');
  if (/<tableParts\b/i.test(xml)) unsupported.push('xlsx:tables');
  if (/<pivotTableDefinition\b/i.test(xml)) unsupported.push('xlsx:pivot-tables');
  if (/<(?:legacyDrawing|oleObjects|controls)\b/i.test(xml)) {
    unsupported.push('xlsx:embedded-objects');
  }
  cells.sort((left, right) => left.row - right.row || left.column - right.column);
  return {
    name,
    id,
    cells,
    merges,
    ...(filter === undefined ? {} : { filter }),
    ...(conditionalFormatting.length === 0 ? {} : { conditionalFormatting }),
    ...(visibility === undefined ? {} : { visibility }),
  };
}

/** Creates an atomic XLSX reader with ZIP and XML resource limits. */
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
        if (/^xl\/drawings\/_rels\//i.test(name)) continue;
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
      const definedNames =
        /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/i.exec(workbook)?.[1] ?? '';
      if (
        [...definedNames.matchAll(/<definedName\b([^>]*?)(?:\/>|>[\s\S]*?<\/definedName>)/gi)].some(
          (match) =>
            !new Set(['_xlnm.Print_Area', '_xlnm.Print_Titles']).has(
              attributes(match[1]!).name ?? '',
            ),
        )
      ) {
        unsupported.push('xlsx:defined-names');
      }
      if (/<workbookProtection\b/i.test(workbook)) unsupported.push('xlsx:workbook-protection');
      const styles = parseStyles(
        entries['xl/styles.xml'] === undefined
          ? undefined
          : archiveXml(entries, 'xl/styles.xml', limits),
      );
      if (entries['xl/calcChain.xml'] !== undefined) unsupported.push('xlsx:calculation-chain');
      if (entryNames.some((name) => /^xl\/comments\d*\.xml$/i.test(name))) {
        unsupported.push('xlsx:comments');
      }
      const sheets: ImportedSheet[] = [];
      const validations: { id: string; value: JsonValue }[] = [];
      const resources: ResourceMetadata[] = [];
      const archiveUncompressedBytes = Object.values(entries).reduce(
        (total, bytes) => total + bytes.byteLength,
        0,
      );
      const drawingResourcePool = createXlsxDrawingResourcePool({
        maxObjects: limits.maxObjects,
        maxResources: limits.maxResources,
        maxResourceBytes: limits.maxResourceBytes,
        maxMaterializedBytes: limits.maxUncompressedBytes - archiveUncompressedBytes,
      });
      const printProfiles: SpreadsheetDocumentInput['templates'][number]['printProfiles'][number][] =
        [];
      let totalCells = 0;
      let sheetIndex = 0;
      for (const match of workbook.matchAll(/<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/gi)) {
        throwIfAborted(options.signal);
        const sheetAttributes = attributes(match[1]!);
        const relationshipId = sheetAttributes['r:id'];
        const target = relationshipId ? relationships.get(relationshipId) : undefined;
        if (!sheetAttributes.name || !target) {
          throw new InterchangeError('MALFORMED_WORKBOOK', 'Worksheet relationship is missing');
        }
        const worksheetName = normalizedWorksheetTarget(target);
        const id = `interchange-sheet-${sheetIndex + 1}`;
        const state = sheetAttributes.state;
        const visibility =
          state === 'hidden'
            ? ('hidden' as const)
            : state === 'veryHidden'
              ? ('very-hidden' as const)
              : ('visible' as const);
        const worksheetXml = archiveXml(entries, worksheetName, limits);
        const parsedSheet = parseWorksheet(
          id,
          sheetAttributes.name,
          worksheetXml,
          strings,
          styles,
          validations,
          limits,
          unsupported,
          visibility,
          options.signal,
        );
        const drawing = parseWorksheetDrawing(
          entries,
          worksheetName,
          worksheetXml,
          id as DocumentSheetId,
          limits,
          drawingResourcePool,
        );
        if (!drawing.unsupported.includes('xlsx:drawing-objects')) {
          const genericDrawing = unsupported.lastIndexOf('xlsx:drawing-objects');
          if (genericDrawing >= 0) unsupported.splice(genericDrawing, 1);
        }
        unsupported.push(...drawing.unsupported);
        resources.push(...drawing.resources);
        const sheet: ImportedSheet = {
          ...parsedSheet,
          ...(drawing.objects.length === 0 ? {} : { objects: drawing.objects }),
        };
        const profile = printProfile(id, worksheetXml);
        if (profile !== undefined) printProfiles.push(profile);
        totalCells += sheet.cells.length;
        if (totalCells > limits.maxCells) {
          throw new InterchangeError('CELL_LIMIT_EXCEEDED', 'XLSX cell limit exceeded');
        }
        sheets.push(sheet);
        sheetIndex += 1;
      }
      if (sheets.length === 0) {
        throw new InterchangeError('MALFORMED_WORKBOOK', 'Workbook contains no worksheets');
      }
      throwIfAborted(options.signal);
      return importResult(
        'xlsx',
        buildDocument(sheets, {
          styles: styles.registry,
          validations,
          resources,
          templates:
            printProfiles.length === 0
              ? []
              : [
                  {
                    id: 'xlsx-imported-print',
                    name: 'Imported print',
                    bindings: [],
                    printProfiles: [...printProfiles],
                  },
                ],
        }),
        [],
        unsupported,
      );
    },
  });
}
