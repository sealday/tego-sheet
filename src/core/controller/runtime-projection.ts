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
  Cell,
  CellInput,
  Sheet,
  SheetInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../document/model/document';
import { parseSpreadsheetDocument } from '../../document/parse-document';

type ValidationId = NonNullable<Cell['validationId']>;

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

  const sheets = document.workbook.sheets.map((sheet): SheetData => {
    const validations = new Map<string, ValidationData>();
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

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function legacyRow(sheet: SheetData | undefined, row: number): Record<string, unknown> | undefined {
  const value = sheet?.rows?.[String(row)];
  return typeof value === 'object' && value !== null
    ? (value as unknown as Record<string, unknown>)
    : undefined;
}

function legacyColumn(
  sheet: SheetData | undefined,
  column: number,
): Record<string, unknown> | undefined {
  const value = sheet?.cols?.[String(column)];
  return typeof value === 'object' && value !== null
    ? (value as unknown as Record<string, unknown>)
    : undefined;
}

function legacyCell(sheet: SheetData | undefined, row: number, column: number): unknown {
  const cells = legacyRow(sheet, row)?.cells;
  return typeof cells === 'object' && cells !== null
    ? (cells as Record<string, unknown>)[String(column)]
    : undefined;
}

function legacyCellInput(cell: unknown): unknown {
  if (typeof cell !== 'object' || cell === null) return cell;
  const record = cell as Record<string, unknown>;
  return {
    ...(record.type === undefined ? {} : { type: record.type }),
    ...(record.text === undefined ? {} : { text: record.text }),
    ...(record.value === undefined ? {} : { value: record.value }),
  };
}

function rowShape(row: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (row === undefined) return undefined;
  const { cells: _cells, ...shape } = row;
  return shape;
}

function mergeSheet(
  previous: Sheet | undefined,
  operational: SheetInput,
  beforeLegacy: SheetData | undefined,
  afterLegacy: SheetData | undefined,
  authoritativeInputs: ReadonlySet<string> | undefined,
  authoritativeValidations: ReadonlyMap<string, ValidationId | null> | undefined,
): SpreadsheetDocumentInput['workbook']['sheets'][number] {
  if (previous === undefined) return operational;
  const previousCells = new Map(
    previous.cells.map((item) => [`${item.row}:${item.column}`, item] as const),
  );
  const operationalCells = new Map(
    operational.cells.map((item) => [`${item.row}:${item.column}`, item] as const),
  );
  const coordinates = new Set([...previousCells.keys(), ...operationalCells.keys()]);
  const cells = [...coordinates]
    .map((key) => {
      const previousCell = previousCells.get(key);
      const operationalCell = operationalCells.get(key);
      const [row, column] = key.split(':').map(Number) as [number, number];
      const beforeCell = legacyCell(beforeLegacy, row, column);
      const afterCell = legacyCell(afterLegacy, row, column);
      if (sameJson(beforeCell, afterCell)) {
        if (authoritativeValidations?.has(key) === true) return previousCell;
        if (
          previousCell !== undefined &&
          operationalCell !== undefined &&
          previousCell.cell.validationId !== operationalCell.cell.validationId
        ) {
          const { validationId: _validationId, ...previousCellData } = previousCell.cell;
          return {
            ...previousCell,
            cell: {
              ...previousCellData,
              ...(operationalCell.cell.validationId === undefined
                ? {}
                : { validationId: operationalCell.cell.validationId }),
            },
          };
        }
        return previousCell ?? operationalCell;
      }
      if (afterCell === undefined || operationalCell === undefined) return undefined;
      const previousCellData = previousCell?.cell;
      const validationId =
        authoritativeValidations?.has(key) === true
          ? (authoritativeValidations.get(key) ?? undefined)
          : operationalCell.cell.validationId;
      return {
        ...operationalCell,
        cell: {
          ...(previousCellData?.resourceId === undefined
            ? {}
            : { resourceId: previousCellData.resourceId }),
          ...(previousCellData?.templateId === undefined
            ? {}
            : { templateId: previousCellData.templateId }),
          ...(previousCellData?.metadata === undefined
            ? {}
            : { metadata: previousCellData.metadata }),
          ...operationalCell.cell,
          ...(validationId === undefined ? {} : { validationId }),
          input:
            previousCellData !== undefined &&
            (authoritativeInputs?.has(key) === true ||
              sameJson(legacyCellInput(beforeCell), legacyCellInput(afterCell)))
              ? previousCellData.input
              : operationalCell.cell.input,
        },
      };
    })
    .filter((item): item is NonNullable<typeof item> => item !== undefined);

  const previousRows = new Map(previous.rows.map((row) => [row.index, row]));
  const operationalRows = new Map((operational.rows ?? []).map((row) => [row.index, row]));
  const rowIndexes = new Set([...previousRows.keys(), ...operationalRows.keys()]);
  const rows = [...rowIndexes]
    .map((index) =>
      sameJson(rowShape(legacyRow(beforeLegacy, index)), rowShape(legacyRow(afterLegacy, index)))
        ? (previousRows.get(index) ?? operationalRows.get(index))
        : operationalRows.get(index),
    )
    .filter((row): row is NonNullable<typeof row> => row !== undefined);

  const previousColumns = new Map(previous.columns.map((column) => [column.index, column]));
  const operationalColumns = new Map(
    (operational.columns ?? []).map((column) => [column.index, column]),
  );
  const columnIndexes = new Set([...previousColumns.keys(), ...operationalColumns.keys()]);
  const columns = [...columnIndexes]
    .map((index) =>
      sameJson(legacyColumn(beforeLegacy, index), legacyColumn(afterLegacy, index))
        ? (previousColumns.get(index) ?? operationalColumns.get(index))
        : operationalColumns.get(index),
    )
    .filter((column): column is NonNullable<typeof column> => column !== undefined);

  const rowCount = sameJson(beforeLegacy?.rows?.len, afterLegacy?.rows?.len)
    ? previous.rowCount
    : operational.rowCount;
  const columnCount = sameJson(beforeLegacy?.cols?.len, afterLegacy?.cols?.len)
    ? previous.columnCount
    : operational.columnCount;
  const freeze = sameJson(beforeLegacy?.freeze, afterLegacy?.freeze)
    ? previous.freeze
    : operational.freeze;
  const filter = sameJson(beforeLegacy?.autofilter, afterLegacy?.autofilter)
    ? previous.filter === undefined
      ? undefined
      : {
          ...(previous.filter.range === undefined
            ? {}
            : {
                range: {
                  start: { ...previous.filter.range.start },
                  end: { ...previous.filter.range.end },
                },
              }),
          filters: previous.filter.filters.map((item) => ({
            ...item,
            values: [...item.values],
          })),
          ...(previous.filter.sort === undefined
            ? {}
            : {
                sort: previous.filter.sort === null ? null : { ...previous.filter.sort },
              }),
        }
    : operational.filter;
  return {
    ...operational,
    name: sameJson(beforeLegacy?.name, afterLegacy?.name) ? previous.name : operational.name,
    cells,
    merges: sameJson(beforeLegacy?.merges, afterLegacy?.merges)
      ? previous.merges.map((range) => ({
          start: { ...range.start },
          end: { ...range.end },
        }))
      : operational.merges,
    ...(rowCount === undefined ? {} : { rowCount }),
    ...(columnCount === undefined ? {} : { columnCount }),
    rows,
    columns,
    ...(freeze === undefined ? {} : { freeze }),
    ...(filter === undefined ? {} : { filter }),
  } as unknown as SheetInput;
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
  beforeWorkbook: WorkbookData,
  afterWorkbook: WorkbookData,
  previous: SpreadsheetDocument,
  sheetIds: readonly SheetId[],
  authoritativeInputs: ReadonlyMap<string, ReadonlySet<string>> = new Map(),
  authoritativeValidations: ReadonlyMap<
    string,
    ReadonlyMap<string, ValidationId | null>
  > = new Map(),
): SpreadsheetDocument {
  const migrated = migrateLegacyWorkbook(afterWorkbook, {
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
      sheets: migrated.document.workbook.sheets.map((sheet, index) =>
        mergeSheet(
          previous.workbook.sheets.find((item) => item.id === sheet.id),
          remapSheet(
            sheet as unknown as SpreadsheetDocumentInput['workbook']['sheets'][number],
            styleRegistry.remap,
            validationRegistry.remap,
          ),
          beforeWorkbook[previous.workbook.sheets.findIndex((item) => item.id === sheet.id)],
          afterWorkbook[index],
          authoritativeInputs.get(sheet.id),
          authoritativeValidations.get(sheet.id),
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
  if (!result.ok) {
    throw new TypeError(
      `Runtime projection produced an invalid spreadsheet document: ${JSON.stringify(result.diagnostics)}`,
    );
  }
  return result.document;
}
