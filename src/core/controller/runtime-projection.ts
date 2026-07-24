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
import type { FormulaValue } from '../../formula/ast';
import { formulaAddressKey } from '../../formula/dependency-graph';
import { parseSpreadsheetDocument } from '../../document/parse-document';
import { legacyValidationRule } from '../../validation/document-rule';

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

function legacyFormulaValue(value: FormulaValue | undefined): CellData['value'] {
  if (value === undefined) return undefined;
  if (value.type === 'blank') return null;
  if (value.type === 'array') return '#SPILL!';
  return value.value;
}

function legacyInput(
  input: CellInput,
  calculated?: FormulaValue,
): Pick<CellData, 'text' | 'value'> & { type?: string } {
  switch (input.type) {
    case 'blank':
      return {};
    case 'string':
      return { text: input.value };
    case 'formula':
      return {
        text: input.source,
        ...(calculated === undefined ? {} : { value: legacyFormulaValue(calculated) }),
      };
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

function normalizeRuntimeWorkbook(workbook: WorkbookData): WorkbookData {
  const normalized = structuredClone(workbook) as SheetData[];
  for (const sheet of normalized) {
    for (const [rowKey, row] of Object.entries(sheet.rows ?? {})) {
      if (rowKey === 'len' || typeof row !== 'object' || row === null || !('cells' in row))
        continue;
      const cells = row.cells;
      if (typeof cells !== 'object' || cells === null) continue;
      for (const cell of Object.values(cells)) {
        if (typeof cell !== 'object' || cell === null) continue;
        const mutable = cell as Record<string, unknown>;
        if (
          (mutable.text === undefined || mutable.text === '') &&
          ((mutable.type === 'number' &&
            (typeof mutable.value !== 'number' || !Number.isFinite(mutable.value))) ||
            (mutable.type === 'boolean' && typeof mutable.value !== 'boolean'))
        ) {
          delete mutable.value;
          delete mutable.type;
        }
      }
    }
  }
  return normalized;
}

/** @internal Projects a schema 2 snapshot into the sole legacy operation boundary. */
export function projectDocumentToLegacy(
  document: SpreadsheetDocument,
  calculatedValues: ReadonlyMap<string, FormulaValue> = new Map(),
): WorkbookData {
  const styles = document.workbook.styles.map((entry) => entry.value as CellStyle);
  const styleIndexes = new Map(document.workbook.styles.map((entry, index) => [entry.id, index]));

  const sheets = document.workbook.sheets.map((sheet): SheetData => {
    const validations = new Map<string, ValidationData>();
    const rows: Record<string, unknown> & { len?: number } = {};
    const cols: Record<string, unknown> & { len?: number } = {};
    if (sheet.rowCount !== undefined) rows.len = sheet.rowCount;
    if (sheet.columnCount !== undefined) cols.len = sheet.columnCount;
    const collapsedRows = new Set<number>();
    const collapsedColumns = new Set<number>();
    for (const group of sheet.groups) {
      if (!group.collapsed) continue;
      const target = group.axis === 'row' ? collapsedRows : collapsedColumns;
      for (let index = group.start; index <= group.end; index += 1) target.add(index);
    }
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
    for (const index of collapsedRows) {
      rows[String(index)] = {
        ...(rows[String(index)] as Record<string, unknown> | undefined),
        hide: true,
      };
    }
    for (const index of collapsedColumns) {
      cols[String(index)] = {
        ...(cols[String(index)] as Record<string, unknown> | undefined),
        hide: true,
      };
    }
    for (const sparse of sheet.cells) {
      const rowKey = String(sparse.row);
      const currentRow = (rows[rowKey] ?? {}) as Record<string, unknown>;
      const cells = (currentRow.cells ?? {}) as Record<string, unknown>;
      cells[String(sparse.column)] = {
        ...legacyInput(
          sparse.cell.input,
          calculatedValues.get(
            formulaAddressKey({
              sheetId: sheet.id,
              row: sparse.row,
              column: sparse.column,
            }),
          ),
        ),
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
          const legacy = {
            ...jsonRecord(registry.value),
            refs: [a1(sparse)],
          } as ValidationData;
          if (legacyValidationRule(legacy) === undefined) continue;
          const existing = validations.get(registry.id);
          validations.set(registry.id, {
            ...legacy,
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

interface IndexInterval {
  readonly start: number;
  readonly end: number;
}

function collapsedGroupIntervals(
  sheet: Sheet,
): Readonly<Record<'row' | 'column', readonly IndexInterval[]>> {
  const intervals: Record<'row' | 'column', IndexInterval[]> = { row: [], column: [] };
  for (const axis of ['row', 'column'] as const) {
    const sorted = sheet.groups
      .filter((group) => group.collapsed && group.axis === axis)
      .map(({ start, end }) => ({ start, end }))
      .sort((left, right) => left.start - right.start || left.end - right.end);
    for (const interval of sorted) {
      const previous = intervals[axis].at(-1);
      if (previous === undefined || interval.start > previous.end + 1) {
        intervals[axis].push(interval);
      } else if (interval.end > previous.end) {
        intervals[axis][intervals[axis].length - 1] = {
          start: previous.start,
          end: interval.end,
        };
      }
    }
  }
  return intervals;
}

function intervalContains(intervals: readonly IndexInterval[], index: number): boolean {
  let lower = 0;
  let upper = intervals.length - 1;
  while (lower <= upper) {
    const middle = lower + Math.floor((upper - lower) / 2);
    const interval = intervals[middle]!;
    if (index < interval.start) upper = middle - 1;
    else if (index > interval.end) lower = middle + 1;
    else return true;
  }
  return false;
}

function stripGroupDerivedHidden<T extends { readonly index: number; readonly hidden?: boolean }>(
  selected: T | undefined,
  explicit: T | undefined,
  derived: boolean,
  hideUnchanged: boolean,
): T | undefined {
  if (selected === undefined || !derived || !hideUnchanged) return selected;
  const { hidden: _derivedHidden, ...layout } = selected;
  const normalized = {
    ...layout,
    ...(explicit?.hidden === undefined ? {} : { hidden: explicit.hidden }),
  } as T;
  return Object.keys(normalized).length === 1 ? undefined : normalized;
}

function mergeSheet(
  previous: Sheet | undefined,
  operational: SheetInput,
  beforeLegacy: SheetData | undefined,
  afterLegacy: SheetData | undefined,
  authoritativeInputs: ReadonlySet<string> | undefined,
  authoritativeValidations: ReadonlyMap<string, ValidationId | null> | undefined,
  legacyValidationIds: ReadonlySet<string>,
): SpreadsheetDocumentInput['workbook']['sheets'][number] {
  if (previous === undefined) return operational;
  const collapsedIntervals = collapsedGroupIntervals(previous);
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
        if (authoritativeValidations?.has(key) !== true) {
          if (
            previousCell !== undefined &&
            operationalCell !== undefined &&
            previousCell.cell.validationId !== operationalCell.cell.validationId &&
            (operationalCell.cell.validationId !== undefined ||
              previousCell.cell.validationId === undefined ||
              legacyValidationIds.has(previousCell.cell.validationId))
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
        if (
          authoritativeValidations.get(key) === null &&
          authoritativeInputs?.has(key) !== true &&
          operationalCell?.cell.input.type === 'blank'
        ) {
          return undefined;
        }
        if (previousCell !== undefined && operationalCell !== undefined) {
          const { validationId: _validationId, ...previousCellData } = previousCell.cell;
          const validationId = authoritativeValidations.get(key) ?? undefined;
          return {
            ...previousCell,
            cell: {
              ...previousCellData,
              ...(validationId === undefined ? {} : { validationId }),
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
          : (previousCellData?.validationId ?? operationalCell.cell.validationId);
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
    .map((index) => {
      const beforeRow = legacyRow(beforeLegacy, index);
      const afterRow = legacyRow(afterLegacy, index);
      const selected = sameJson(rowShape(beforeRow), rowShape(afterRow))
        ? (previousRows.get(index) ?? operationalRows.get(index))
        : operationalRows.get(index);
      return stripGroupDerivedHidden(
        selected,
        previousRows.get(index),
        intervalContains(collapsedIntervals.row, index),
        beforeRow?.hide === afterRow?.hide,
      );
    })
    .filter((row): row is NonNullable<typeof row> => row !== undefined);

  const previousColumns = new Map(previous.columns.map((column) => [column.index, column]));
  const operationalColumns = new Map(
    (operational.columns ?? []).map((column) => [column.index, column]),
  );
  const columnIndexes = new Set([...previousColumns.keys(), ...operationalColumns.keys()]);
  const columns = [...columnIndexes]
    .map((index) => {
      const beforeColumn = legacyColumn(beforeLegacy, index);
      const afterColumn = legacyColumn(afterLegacy, index);
      const selected = sameJson(beforeColumn, afterColumn)
        ? (previousColumns.get(index) ?? operationalColumns.get(index))
        : operationalColumns.get(index);
      return stripGroupDerivedHidden(
        selected,
        previousColumns.get(index),
        intervalContains(collapsedIntervals.column, index),
        beforeColumn?.hide === afterColumn?.hide,
      );
    })
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
    visibility: previous.visibility,
    conditionalFormatting: previous.conditionalFormatting.map((format) => structuredClone(format)),
    filterViews: previous.filterViews.map((view) => structuredClone(view)),
    objects: previous.objects.map((object) => structuredClone(object)),
    tables: previous.tables.map((table) => structuredClone(table)),
    groups: previous.groups.map((group) => ({ ...group })),
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
  const migrated = migrateLegacyWorkbook(normalizeRuntimeWorkbook(afterWorkbook), {
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
  const legacyValidationIds = new Set(
    previous.workbook.validations
      .filter(({ value }) => legacyValidationRule({ ...jsonRecord(value), refs: [] }) !== undefined)
      .map(({ id }) => id),
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
          legacyValidationIds,
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
