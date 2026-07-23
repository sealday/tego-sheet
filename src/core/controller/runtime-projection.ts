import type {
  AutoFilterData,
  CellData,
  CellStyle,
  ColsData,
  JsonValue,
  RowsData,
  SheetData,
  ValidationData,
  WorkbookData,
} from '..';
import type { SheetId } from '../types/coordinates';
import { migrateLegacyWorkbook } from '../../document/migrate-legacy';
import type {
  CellInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../document/model/document';
import { parseSpreadsheetDocument } from '../../document/parse-document';

function columnName(column: number): string {
  let value = column + 1;
  let output = '';
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + (value % 26)) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

function a1(point: { readonly row: number; readonly column: number }): string {
  return `${columnName(point.column)}${point.row + 1}`;
}

function rangeA1(range: {
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
}): string {
  const start = a1(range.start);
  const end = a1(range.end);
  return start === end ? start : `${start}:${end}`;
}

function legacyInput(input: CellInput): Pick<CellData, 'text' | 'value'> & { type?: string } {
  switch (input.type) {
    case 'blank':
      return {};
    case 'string':
      return { text: input.value };
    case 'formula':
      return { text: input.source };
    case 'number':
      return { type: 'number', value: input.value };
    case 'boolean':
      return { type: 'boolean', value: input.value };
    case 'custom':
      return { value: input.value };
  }
}

function jsonRecord(value: JsonValue): Record<string, JsonValue> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, JsonValue>)
    : {};
}

/** @internal Projects a schema 2 snapshot into the sole legacy operation boundary. */
export function projectDocumentToLegacy(document: SpreadsheetDocument): WorkbookData {
  const styles = document.workbook.styles.map((entry) => entry.value as CellStyle);
  const styleIndexes = new Map(document.workbook.styles.map((entry, index) => [entry.id, index]));
  const validations = new Map<string, ValidationData>();

  const sheets = document.workbook.sheets.map((sheet): SheetData => {
    const rows: Record<string, unknown> & { len?: number } = {};
    const cols: Record<string, unknown> & { len?: number } = {};
    if (sheet.rowCount !== undefined) rows.len = sheet.rowCount;
    if (sheet.columnCount !== undefined) cols.len = sheet.columnCount;
    for (const row of sheet.rows) {
      rows[String(row.index)] = {
        ...(row.height === undefined ? {} : { height: row.height }),
        ...(row.hidden === undefined ? {} : { hide: row.hidden }),
        ...(row.styleId === undefined ? {} : { style: styleIndexes.get(row.styleId) }),
      };
    }
    for (const column of sheet.columns) {
      cols[String(column.index)] = {
        ...(column.width === undefined ? {} : { width: column.width }),
        ...(column.hidden === undefined ? {} : { hide: column.hidden }),
        ...(column.styleId === undefined ? {} : { style: styleIndexes.get(column.styleId) }),
      };
    }
    for (const sparse of sheet.cells) {
      const rowKey = String(sparse.row);
      const currentRow = (rows[rowKey] ?? {}) as Record<string, unknown>;
      const cells = (currentRow.cells ?? {}) as Record<string, unknown>;
      cells[String(sparse.column)] = {
        ...legacyInput(sparse.cell.input),
        ...(sparse.cell.styleId === undefined
          ? {}
          : { style: styleIndexes.get(sparse.cell.styleId) }),
        ...(sparse.cell.editable === undefined ? {} : { editable: sparse.cell.editable }),
        ...(sparse.cell.printable === undefined ? {} : { printable: sparse.cell.printable }),
      };
      rows[rowKey] = { ...currentRow, cells };
      if (sparse.cell.validationId !== undefined) {
        const registry = document.workbook.validations.find(
          (entry) => entry.id === sparse.cell.validationId,
        );
        if (registry !== undefined) {
          const existing = validations.get(registry.id);
          validations.set(registry.id, {
            ...jsonRecord(registry.value),
            refs: [...(existing?.refs ?? []), a1(sparse)],
          });
        }
      }
    }
    const filter: AutoFilterData | undefined =
      sheet.filter === undefined
        ? undefined
        : {
            ...(sheet.filter.range === undefined ? {} : { ref: rangeA1(sheet.filter.range) }),
            filters: sheet.filter.filters.map((item) => ({
              ci: item.column,
              operator: item.operator,
              value: item.values,
            })),
            ...(sheet.filter.sort === undefined
              ? {}
              : {
                  sort:
                    sheet.filter.sort === null
                      ? null
                      : {
                          ci: sheet.filter.sort.column,
                          order: sheet.filter.sort.direction,
                        },
                }),
          };
    return {
      name: sheet.name,
      styles,
      ...(Object.keys(rows).length === 0 ? {} : { rows: rows as RowsData }),
      ...(Object.keys(cols).length === 0 ? {} : { cols: cols as ColsData }),
      merges: sheet.merges.map(rangeA1),
      validations: [...validations.values()],
      ...(sheet.freeze === undefined ? {} : { freeze: a1(sheet.freeze) }),
      ...(filter === undefined ? {} : { autofilter: filter }),
    } as unknown as SheetData;
  });
  return sheets;
}

function registryIds(
  previous: readonly { readonly id: string; readonly value: JsonValue }[],
  next: readonly { readonly id: string; readonly value: JsonValue }[],
  prefix: string,
): {
  readonly entries: { id: string; value: JsonValue }[];
  readonly remap: ReadonlyMap<string, string>;
} {
  const byValue = new Map(previous.map((entry) => [JSON.stringify(entry.value), entry.id]));
  const used = new Set<string>();
  const remap = new Map<string, string>();
  const entries = next.map((entry, index) => {
    let id = byValue.get(JSON.stringify(entry.value)) ?? `${prefix}-${index + 1}`;
    while (used.has(id)) id = `${prefix}-${index + used.size + 1}`;
    used.add(id);
    remap.set(entry.id, id);
    return { id, value: entry.value };
  });
  for (const entry of previous) {
    if (!used.has(entry.id)) entries.push({ id: entry.id, value: entry.value });
  }
  return { entries, remap };
}

function remapSheet(
  sheet: SpreadsheetDocumentInput['workbook']['sheets'][number],
  styles: ReadonlyMap<string, string>,
  validations: ReadonlyMap<string, string>,
): SpreadsheetDocumentInput['workbook']['sheets'][number] {
  return {
    ...sheet,
    cells: sheet.cells.map((item) => ({
      ...item,
      cell: {
        ...item.cell,
        ...(item.cell.styleId === undefined
          ? {}
          : { styleId: styles.get(item.cell.styleId) ?? item.cell.styleId }),
        ...(item.cell.validationId === undefined
          ? {}
          : {
              validationId: validations.get(item.cell.validationId) ?? item.cell.validationId,
            }),
      },
    })),
    rows: sheet.rows?.map((row) => ({
      ...row,
      ...(row.styleId === undefined ? {} : { styleId: styles.get(row.styleId) ?? row.styleId }),
    })),
    columns: sheet.columns?.map((column) => ({
      ...column,
      ...(column.styleId === undefined
        ? {}
        : { styleId: styles.get(column.styleId) ?? column.styleId }),
    })),
  };
}

/** @internal Rebuilds the frozen schema 2 truth after one legacy operation commit. */
export function projectLegacyToDocument(
  workbook: WorkbookData,
  previous: SpreadsheetDocument,
  sheetIds: readonly SheetId[],
): SpreadsheetDocument {
  const migrated = migrateLegacyWorkbook(workbook, {
    ids: {
      documentId: () => previous.id,
      sheetId: (index) => sheetIds[index] ?? `runtime-sheet-${index + 1}`,
    },
    dateSystem: previous.workbook.settings.dateSystem,
    localeHint: previous.workbook.settings.localeHint,
  });
  if (!migrated.ok)
    throw new TypeError('Runtime workbook could not produce a spreadsheet document');
  const styleRegistry = registryIds(
    previous.workbook.styles,
    migrated.document.workbook.styles,
    'runtime-style',
  );
  const validationRegistry = registryIds(
    previous.workbook.validations,
    migrated.document.workbook.validations,
    'runtime-validation',
  );
  const result = parseSpreadsheetDocument({
    ...migrated.document,
    workbook: {
      ...migrated.document.workbook,
      sheets: migrated.document.workbook.sheets.map((sheet) =>
        remapSheet(
          sheet as unknown as SpreadsheetDocumentInput['workbook']['sheets'][number],
          styleRegistry.remap,
          validationRegistry.remap,
        ),
      ),
      styles: styleRegistry.entries,
      validations: validationRegistry.entries,
      settings: previous.workbook.settings,
    },
    templates: previous.templates,
    resources: previous.resources,
    extensions: previous.extensions,
  });
  if (!result.ok)
    throw new TypeError('Runtime projection produced an invalid spreadsheet document');
  return result.document;
}
