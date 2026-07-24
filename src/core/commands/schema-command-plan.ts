import type { WorkbookCommand } from './workbook-command';
import {
  assertClipboardResourceLimit,
  internalPasteRange,
  pasteInternal,
} from '../operations/clipboard';
import type { CellRange, SheetId } from '../types/coordinates';
import type { WorkbookChangeKind } from '../types/changes';
import type { WorkbookData } from '../types/workbook';
import type {
  Cell,
  SheetInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../document/model/document';
import { parseSpreadsheetDocument } from '../../document/parse-document';
import {
  CoordinateTransform,
  transformDocumentCoordinates,
} from '../coordinates/coordinate-transform';
import { parseFormula, renameFormulaSheet, renderFormula } from '../../formula';
import { canonicalAutofillTargetRange, createTypedAutofillResolver } from './typed-autofill';

type ValidationId = NonNullable<Cell['validationId']>;
type MutableSheetGroup = NonNullable<SheetInput['groups']>[number];

export class GroupLimitExceededError extends Error {
  readonly code = 'GROUP_LIMIT_EXCEEDED';
}

export interface SchemaCommandPlan {
  readonly document: SpreadsheetDocument;
  readonly authoritativeInputs: ReadonlyMap<string, ReadonlySet<string>>;
  readonly authoritativeValidations: ReadonlyMap<string, ReadonlyMap<string, ValidationId | null>>;
}

export interface SchemaProjectionCommit {
  readonly result: unknown;
  readonly kind: Extract<
    WorkbookChangeKind,
    'cell' | 'clipboard' | 'autofill' | 'view' | 'object' | 'style' | 'outline' | 'structure'
  >;
  readonly sheet: SheetId;
  readonly range?: CellRange;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function sheetIndex(sheetIds: readonly SheetId[], sheet: SheetId): number {
  return sheetIds.findIndex((candidate) => candidate === sheet);
}

/** Pure target range shared by command planning and validation boundaries. */
export function plannedPasteTargetRange(
  command: Extract<WorkbookCommand, { readonly type: 'paste-internal' | 'autofill' }>,
): CellRange {
  if (command.type === 'autofill') {
    return canonicalAutofillTargetRange(command.source.range, command.target.range);
  }
  return internalPasteRange(command.source.range, command.target.range, command.cut);
}

function normalizeOutlineGroups(groups: readonly MutableSheetGroup[]): MutableSheetGroup[] {
  const output: MutableSheetGroup[] = [];
  for (const axis of ['row', 'column'] as const) {
    const axisGroups = groups
      .filter((group) => group.axis === axis)
      .sort(
        (left, right) =>
          left.start - right.start || right.end - left.end || left.id.localeCompare(right.id),
      );
    const stack: MutableSheetGroup[] = [];
    for (const group of axisGroups) {
      while (stack.length > 0 && group.start > stack.at(-1)!.end) stack.pop();
      while (stack.length > 0 && group.end > stack.at(-1)!.end) stack.pop();
      const normalized = { ...group, level: stack.length + 1 };
      output.push(normalized);
      stack.push(normalized);
    }
  }
  return output;
}

function transformOutlineGroups(
  groups: readonly MutableSheetGroup[],
  command: Extract<
    WorkbookCommand,
    { readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' }
  >,
): MutableSheetGroup[] {
  const axis = command.type.endsWith('row') ? 'row' : 'column';
  const count = command.count ?? 1;
  const deletionEnd = command.index + count - 1;
  const output: MutableSheetGroup[] = [];
  for (const group of groups) {
    if (group.axis !== axis) {
      output.push({ ...group });
      continue;
    }
    if (command.type === 'insert-row' || command.type === 'insert-column') {
      output.push({
        ...group,
        ...(command.index <= group.start
          ? { start: group.start + count, end: group.end + count }
          : command.index <= group.end
            ? { end: group.end + count }
            : {}),
      });
      continue;
    }
    if (group.end < command.index) {
      output.push({ ...group });
      continue;
    }
    if (group.start > deletionEnd) {
      output.push({ ...group, start: group.start - count, end: group.end - count });
      continue;
    }
    const beforeEnd = Math.min(group.end, command.index - 1);
    const afterStart = Math.max(group.start, deletionEnd + 1);
    const hasBefore = group.start <= beforeEnd;
    const hasAfter = afterStart <= group.end;
    if (!hasBefore && !hasAfter) continue;
    output.push({
      ...group,
      start: hasBefore ? group.start : afterStart - count,
      end: hasAfter ? group.end - count : beforeEnd,
    });
  }
  return normalizeOutlineGroups(output);
}

function transformStructure(
  input: SpreadsheetDocumentInput,
  command: Extract<
    WorkbookCommand,
    { readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' }
  >,
  sheetIds: readonly SheetId[],
  authoritativeInputs: Map<string, Set<string>>,
  authoritativeValidations: Map<string, Map<string, ValidationId | null>>,
): void {
  const index = sheetIndex(sheetIds, command.sheet);
  const sheet = input.workbook.sheets[index];
  if (sheet === undefined) return;
  const axis = command.type.endsWith('row') ? 'row' : 'column';
  const transform =
    command.type === 'insert-row' || command.type === 'insert-column'
      ? CoordinateTransform.insert(axis, command.index, command.count ?? 1)
      : CoordinateTransform.delete(axis, command.index, command.count ?? 1);
  sheet.groups = transformOutlineGroups(sheet.groups ?? [], command);
  const transformedDocument = transformDocumentCoordinates(input, sheet.id, transform);
  input.workbook = transformedDocument.workbook;
  input.templates = transformedDocument.templates;
  const transformedSheet = input.workbook.sheets[index] as SheetInput;
  for (const candidate of input.workbook.sheets) {
    const formulaKeys = new Set(
      candidate.cells
        .filter((item) => item.cell.input.type === 'formula')
        .map((item) => cellKey(item.row, item.column)),
    );
    if (formulaKeys.size > 0) authoritativeInputs.set(candidate.id, formulaKeys);
  }
  authoritativeValidations.set(
    transformedSheet.id,
    new Map(
      transformedSheet.cells
        .filter((item) => item.cell.validationId !== undefined)
        .map((item) => [cellKey(item.row, item.column), item.cell.validationId as ValidationId]),
    ),
  );
}

function transformSheetRename(
  input: SpreadsheetDocumentInput,
  command: Extract<WorkbookCommand, { readonly type: 'rename-sheet' }>,
  sheetIds: readonly SheetId[],
  authoritativeInputs: Map<string, Set<string>>,
): void {
  const renamedSheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (renamedSheet === undefined || renamedSheet.name === command.name) return;
  const previousName = renamedSheet.name;
  renamedSheet.name = command.name;
  for (const sheet of input.workbook.sheets) {
    const formulaKeys = new Set<string>();
    for (const item of sheet.cells) {
      if (item.cell.input.type !== 'formula') continue;
      let source: string;
      try {
        source = renderFormula(
          renameFormulaSheet(parseFormula(item.cell.input.source), previousName, command.name),
        );
      } catch {
        source = renameLegacySheetQualifiers(item.cell.input.source, previousName, command.name);
      }
      if (source === item.cell.input.source) continue;
      item.cell = { ...item.cell, input: { ...item.cell.input, source } };
      formulaKeys.add(cellKey(item.row, item.column));
    }
    sheet.conditionalFormatting = (sheet.conditionalFormatting ?? []).map((format) => {
      if (format.type === 'color-scale') return format;
      const rename = (formula: string): string => {
        const source = `=${formula}`;
        try {
          return renderFormula(
            renameFormulaSheet(parseFormula(source), previousName, command.name),
          ).slice(1);
        } catch {
          return renameLegacySheetQualifiers(source, previousName, command.name).slice(1);
        }
      };
      return {
        ...format,
        formula: rename(format.formula),
        ...(format.formula2 === undefined ? {} : { formula2: rename(format.formula2) }),
      };
    });
    if (formulaKeys.size > 0) authoritativeInputs.set(sheet.id, formulaKeys);
  }
}

function renderSheetQualifier(name: string): string {
  return /^[A-Z_][A-Z0-9_.]*$/iu.test(name) ? `${name}!` : `'${name.replaceAll("'", "''")}'!`;
}

function renameLegacySheetQualifiers(
  source: string,
  previousName: string,
  nextName: string,
): string {
  let output = '';
  let index = 0;
  let inString = false;
  while (index < source.length) {
    const character = source[index] as string;
    if (character === '"') {
      output += character;
      if (inString && source[index + 1] === '"') {
        output += '"';
        index += 2;
        continue;
      }
      inString = !inString;
      index += 1;
      continue;
    }
    if (!inString && character === "'") {
      let cursor = index + 1;
      let sheetName = '';
      while (cursor < source.length) {
        if (source[cursor] !== "'") {
          sheetName += source[cursor];
          cursor += 1;
          continue;
        }
        if (source[cursor + 1] === "'") {
          sheetName += "'";
          cursor += 2;
          continue;
        }
        break;
      }
      if (source[cursor] === "'" && source[cursor + 1] === '!') {
        output +=
          sheetName.toLowerCase() === previousName.toLowerCase()
            ? renderSheetQualifier(nextName)
            : source.slice(index, cursor + 2);
        index = cursor + 2;
        continue;
      }
    }
    if (!inString) {
      const qualifier = /^[A-Z_][A-Z0-9_.]*!/iu.exec(source.slice(index));
      if (qualifier !== null) {
        const sheetName = qualifier[0].slice(0, -1);
        output +=
          sheetName.toLowerCase() === previousName.toLowerCase()
            ? renderSheetQualifier(nextName)
            : qualifier[0];
        index += qualifier[0].length;
        continue;
      }
    }
    output += character;
    index += 1;
  }
  return output;
}

function getCell(sheet: SheetInput, row: number, column: number): Cell | undefined {
  return sheet.cells.find((item) => item.row === row && item.column === column)?.cell as
    | Cell
    | undefined;
}

function setCell(sheet: SheetInput, row: number, column: number, cell: Cell | undefined): void {
  const index = sheet.cells.findIndex((item) => item.row === row && item.column === column);
  if (cell === undefined) {
    if (index >= 0) sheet.cells.splice(index, 1);
    return;
  }
  const next = { row, column, cell } as SheetInput['cells'][number];
  if (index < 0) sheet.cells.push(next);
  else sheet.cells[index] = next;
}

function setFilterView(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'set-filter-view' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const views = [...(sheet.filterViews ?? [])];
  const index = views.findIndex((view) => view.id === command.view.id);
  if (index < 0) views.push(structuredClone(command.view));
  else views[index] = structuredClone(command.view);
  sheet.filterViews = views;
}

function removeFilterView(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'remove-filter-view' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  sheet.filterViews = (sheet.filterViews ?? []).filter((view) => view.id !== command.viewId);
}

function setConditionalFormat(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'set-conditional-format' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const formats = [...(sheet.conditionalFormatting ?? [])];
  if (command.index > formats.length) {
    throw new RangeError('conditional format index is outside the rule list');
  }
  if (command.index === formats.length) formats.push(structuredClone(command.format));
  else formats[command.index] = structuredClone(command.format);
  sheet.conditionalFormatting = formats;
}

function removeConditionalFormat(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'remove-conditional-format' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const formats = [...(sheet.conditionalFormatting ?? [])];
  if (command.index >= formats.length) {
    throw new RangeError('conditional format index is outside the rule list');
  }
  formats.splice(command.index, 1);
  sheet.conditionalFormatting = formats;
}

function setSheetObject(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'set-sheet-object' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const objects = [...(sheet.objects ?? [])];
  const index = objects.findIndex((object) => object.id === command.object.id);
  if (index < 0) objects.push(structuredClone(command.object));
  else objects[index] = structuredClone(command.object);
  sheet.objects = objects;
}

function removeSheetObject(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'remove-sheet-object' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  sheet.objects = (sheet.objects ?? []).filter((object) => object.id !== command.objectId);
}

function setStructuredTable(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'set-table' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const tables = [...(sheet.tables ?? [])];
  const index = tables.findIndex((table) => table.id === command.table.id);
  const snapshot = {
    ...structuredClone(command.table),
    columns: command.table.columns.map((column) => ({ ...column })),
  };
  if (index < 0) tables.push(snapshot);
  else tables[index] = snapshot;
  sheet.tables = tables;
}

function removeStructuredTable(
  input: SpreadsheetDocumentInput,
  sheetIds: readonly SheetId[],
  command: Extract<WorkbookCommand, { readonly type: 'remove-table' }>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  sheet.tables = (sheet.tables ?? []).filter((table) => table.id !== command.tableId);
}

function mapPasteCell(
  target: Cell | undefined,
  source: Cell | undefined,
  mode: 'all' | 'value' | 'format',
): Cell | undefined {
  if (mode === 'all') return source === undefined ? undefined : structuredClone(source);
  if (mode === 'value') {
    if (source === undefined) return target;
    return {
      ...target,
      input: structuredClone(source.input),
    } as Cell;
  }
  if (target === undefined && source?.styleId === undefined) return undefined;
  const { styleId: _styleId, ...targetWithoutStyle } = target ?? {
    input: { type: 'blank' as const },
  };
  return {
    ...targetWithoutStyle,
    ...(source?.styleId === undefined ? {} : { styleId: source.styleId }),
  };
}

function transformInternalPaste(
  input: SpreadsheetDocumentInput,
  command: Extract<WorkbookCommand, { readonly type: 'paste-internal' | 'autofill' }>,
  sheetIds: readonly SheetId[],
  authoritativeInputs: Map<string, Set<string>>,
  authoritativeValidations: Map<string, Map<string, ValidationId | null>>,
): void {
  const sourceIndex = sheetIndex(sheetIds, command.source.sheet);
  const targetIndex = sheetIndex(sheetIds, command.target.sheet);
  const sourceSheet = input.workbook.sheets[sourceIndex];
  const targetSheet = input.workbook.sheets[targetIndex];
  if (sourceSheet === undefined || targetSheet === undefined) return;
  const source = command.source.range;
  const range = plannedPasteTargetRange(command);
  assertClipboardResourceLimit(source);
  assertClipboardResourceLimit(range);
  const sourceRows = source.end.row - source.start.row + 1;
  const sourceColumns = source.end.column - source.start.column + 1;
  const snapshots = new Map<string, Cell | undefined>();
  for (let row = source.start.row; row <= source.end.row; row += 1) {
    for (let column = source.start.column; column <= source.end.column; column += 1) {
      snapshots.set(cellKey(row, column), structuredClone(getCell(sourceSheet, row, column)));
    }
  }
  const autofillInput =
    command.type === 'autofill'
      ? createTypedAutofillResolver(source, range, (row, column) =>
          snapshots.get(cellKey(row, column)),
        )
      : undefined;
  if (
    command.type === 'paste-internal' &&
    command.cut &&
    sourceIndex === targetIndex &&
    command.mode === 'all'
  ) {
    for (let row = source.start.row; row <= source.end.row; row += 1) {
      for (let column = source.start.column; column <= source.end.column; column += 1) {
        setCell(sourceSheet, row, column, undefined);
        const validations = authoritativeValidations.get(sourceSheet.id) ?? new Map();
        validations.set(cellKey(row, column), null);
        authoritativeValidations.set(sourceSheet.id, validations);
      }
    }
  }
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let column = range.start.column; column <= range.end.column; column += 1) {
      const sourceRow = source.start.row + ((row - range.start.row) % sourceRows);
      const sourceColumn = source.start.column + ((column - range.start.column) % sourceColumns);
      const sourceCell = snapshots.get(cellKey(sourceRow, sourceColumn));
      const targetCell = getCell(targetSheet, row, column);
      if (command.type === 'paste-internal' && command.cut && sourceCell === undefined) continue;
      const mapped = mapPasteCell(targetCell, sourceCell, command.mode);
      const filledInput =
        autofillInput !== undefined && command.mode !== 'format'
          ? autofillInput(row, column)
          : undefined;
      setCell(
        targetSheet,
        row,
        column,
        mapped === undefined || filledInput === undefined
          ? mapped
          : { ...mapped, input: filledInput },
      );
      if (
        command.mode !== 'format' &&
        (command.type === 'autofill' || sourceCell?.input.type === 'custom')
      ) {
        const keys = authoritativeInputs.get(targetSheet.id) ?? new Set<string>();
        keys.add(cellKey(row, column));
        authoritativeInputs.set(targetSheet.id, keys);
      }
      if (command.mode === 'all') {
        const validations = authoritativeValidations.get(targetSheet.id) ?? new Map();
        validations.set(cellKey(row, column), sourceCell?.validationId ?? null);
        authoritativeValidations.set(targetSheet.id, validations);
      }
    }
  }
}

function setTypedCellInput(
  input: SpreadsheetDocumentInput,
  command: Extract<WorkbookCommand, { readonly type: 'set-cell-input' }>,
  sheetIds: readonly SheetId[],
  authoritativeInputs: Map<string, Set<string>>,
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.address.sheet)];
  if (sheet === undefined) return;
  const current = getCell(sheet, command.address.row, command.address.column);
  setCell(sheet, command.address.row, command.address.column, {
    ...current,
    input: structuredClone(command.input),
  });
  const keys = authoritativeInputs.get(sheet.id) ?? new Set<string>();
  keys.add(cellKey(command.address.row, command.address.column));
  authoritativeInputs.set(sheet.id, keys);
}

function mutateOutlineGroups(
  input: SpreadsheetDocumentInput,
  command: Extract<WorkbookCommand, { readonly type: 'group' | 'ungroup' | 'toggle-group' }>,
  sheetIds: readonly SheetId[],
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const groups = [...(sheet.groups ?? [])];
  if (command.type === 'group') {
    if (groups.some((group) => group.id === command.group.id)) {
      throw new GroupLimitExceededError(`Duplicate outline group ID ${command.group.id}`);
    }
    groups.push({ ...command.group, level: 1 });
  } else {
    const index = groups.findIndex((group) => group.id === command.id);
    if (index < 0) {
      throw new GroupLimitExceededError(`Unknown outline group ID ${command.id}`);
    }
    if (command.type === 'ungroup') groups.splice(index, 1);
    else groups[index] = { ...groups[index]!, collapsed: !groups[index]!.collapsed };
  }
  sheet.groups = normalizeOutlineGroups(groups);
}

function mutateExplicitHidden(
  input: SpreadsheetDocumentInput,
  command: Extract<WorkbookCommand, { readonly type: 'set-row-hidden' | 'set-column-hidden' }>,
  sheetIds: readonly SheetId[],
): void {
  const sheet = input.workbook.sheets[sheetIndex(sheetIds, command.sheet)];
  if (sheet === undefined) return;
  const count = command.count ?? 1;
  if (command.type === 'set-row-hidden') {
    const rows = new Map((sheet.rows ?? []).map((row) => [row.index, row]));
    for (let offset = 0; offset < count; offset += 1) {
      const index = command.row + offset;
      rows.set(index, { ...rows.get(index), index, hidden: command.hidden });
    }
    sheet.rows = [...rows.values()].sort((left, right) => left.index - right.index);
    return;
  }
  const columns = new Map((sheet.columns ?? []).map((column) => [column.index, column]));
  for (let offset = 0; offset < count; offset += 1) {
    const index = command.column + offset;
    columns.set(index, { ...columns.get(index), index, hidden: command.hidden });
  }
  sheet.columns = [...columns.values()].sort((left, right) => left.index - right.index);
}

export function prepareSchemaProjectionCommit(
  command: Extract<
    WorkbookCommand,
    {
      readonly type:
        | 'paste-internal'
        | 'autofill'
        | 'set-filter-view'
        | 'remove-filter-view'
        | 'set-conditional-format'
        | 'remove-conditional-format'
        | 'set-sheet-object'
        | 'remove-sheet-object'
        | 'set-table'
        | 'remove-table'
        | 'set-cell-input'
        | 'group'
        | 'ungroup'
        | 'toggle-group'
        | 'set-row-hidden'
        | 'set-column-hidden';
    }
  >,
  projection: WorkbookData,
  sheetIds: readonly SheetId[],
  capturePasteValues: boolean,
): SchemaProjectionCommit {
  if (command.type === 'set-cell-input') {
    const point = { row: command.address.row, column: command.address.column };
    return {
      result: undefined,
      kind: 'cell',
      sheet: command.address.sheet,
      range: { start: point, end: point },
    };
  }
  if (command.type === 'group' || command.type === 'ungroup' || command.type === 'toggle-group') {
    return { result: undefined, kind: 'outline', sheet: command.sheet };
  }
  if (command.type === 'set-row-hidden' || command.type === 'set-column-hidden') {
    return { result: undefined, kind: 'structure', sheet: command.sheet };
  }
  if (command.type === 'set-filter-view' || command.type === 'remove-filter-view') {
    return { result: undefined, kind: 'view', sheet: command.sheet };
  }
  if (command.type === 'set-conditional-format' || command.type === 'remove-conditional-format') {
    return { result: undefined, kind: 'style', sheet: command.sheet };
  }
  if (command.type === 'set-sheet-object' || command.type === 'remove-sheet-object') {
    return { result: undefined, kind: 'object', sheet: command.sheet };
  }
  if (command.type === 'set-table' || command.type === 'remove-table') {
    return {
      result: undefined,
      kind: 'structure',
      sheet: command.sheet,
      ...(command.type === 'remove-table'
        ? {}
        : { range: { start: command.table.range.start, end: command.table.range.end } }),
    };
  }
  const range = plannedPasteTargetRange(command);
  if (command.type === 'autofill') {
    return {
      result: undefined,
      kind: 'autofill',
      sheet: command.target.sheet,
      range,
    };
  }
  const sourceSheet = projection[sheetIndex(sheetIds, command.source.sheet)];
  const targetSheet = projection[sheetIndex(sheetIds, command.target.sheet)];
  if (sourceSheet === undefined || targetSheet === undefined) {
    throw new Error('Schema-only clipboard projection lost its source or target sheet');
  }
  return {
    result: pasteInternal(
      targetSheet,
      sourceSheet,
      command.source.range,
      command.target.range,
      command.mode,
      command.cut,
      capturePasteValues,
    ).values,
    kind: 'clipboard',
    sheet: command.target.sheet,
    range,
  };
}

export function prepareSchemaCommand(
  document: SpreadsheetDocument,
  command: WorkbookCommand,
  sheetIds: readonly SheetId[],
): SchemaCommandPlan {
  const input = structuredClone(document) as unknown as SpreadsheetDocumentInput;
  const authoritativeInputs = new Map<string, Set<string>>();
  const authoritativeValidations = new Map<string, Map<string, ValidationId | null>>();
  switch (command.type) {
    case 'group':
    case 'ungroup':
    case 'toggle-group':
      mutateOutlineGroups(input, command, sheetIds);
      break;
    case 'set-row-hidden':
    case 'set-column-hidden':
      mutateExplicitHidden(input, command, sheetIds);
      break;
    case 'set-cell-input':
      setTypedCellInput(input, command, sheetIds, authoritativeInputs);
      break;
    case 'insert-row':
    case 'delete-row':
    case 'insert-column':
    case 'delete-column':
      transformStructure(input, command, sheetIds, authoritativeInputs, authoritativeValidations);
      break;
    case 'rename-sheet':
      transformSheetRename(input, command, sheetIds, authoritativeInputs);
      break;
    case 'paste-internal':
    case 'autofill':
      transformInternalPaste(
        input,
        command,
        sheetIds,
        authoritativeInputs,
        authoritativeValidations,
      );
      break;
    case 'set-filter-view':
      setFilterView(input, sheetIds, command);
      break;
    case 'remove-filter-view':
      removeFilterView(input, sheetIds, command);
      break;
    case 'set-conditional-format':
      setConditionalFormat(input, sheetIds, command);
      break;
    case 'remove-conditional-format':
      removeConditionalFormat(input, sheetIds, command);
      break;
    case 'set-sheet-object':
      setSheetObject(input, sheetIds, command);
      break;
    case 'remove-sheet-object':
      removeSheetObject(input, sheetIds, command);
      break;
    case 'set-table':
      setStructuredTable(input, sheetIds, command);
      break;
    case 'remove-table':
      removeStructuredTable(input, sheetIds, command);
      break;
  }
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) {
    const groupFailure = parsed.diagnostics.find(
      ({ code }) => code === 'GROUP_LIMIT_EXCEEDED' || code === 'DUPLICATE_ID',
    );
    if (groupFailure !== undefined) throw new GroupLimitExceededError(groupFailure.message);
    throw new TypeError(
      `Schema command plan produced an invalid spreadsheet document: ${JSON.stringify(parsed.diagnostics)}`,
    );
  }
  return { document: parsed.document, authoritativeInputs, authoritativeValidations };
}
