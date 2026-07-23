import type { WorkbookCommand } from './workbook-command';
import { assertClipboardResourceLimit, internalPasteRange } from '../operations/clipboard';
import type { CellRange, SheetId } from '../types/coordinates';
import type {
  Cell,
  SheetInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../document/model/document';
import { parseSpreadsheetDocument } from '../../document/parse-document';

type ValidationId = NonNullable<Cell['validationId']>;

export interface SchemaCommandPlan {
  readonly document: SpreadsheetDocument;
  readonly authoritativeInputs: ReadonlyMap<string, ReadonlySet<string>>;
  readonly authoritativeValidations: ReadonlyMap<string, ReadonlyMap<string, ValidationId | null>>;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function sheetIndex(sheetIds: readonly SheetId[], sheet: SheetId): number {
  return sheetIds.findIndex((candidate) => candidate === sheet);
}

function transformIndex(
  index: number,
  command: Extract<
    WorkbookCommand,
    { readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' }
  >,
): number | null {
  const count = command.count ?? 1;
  if (command.type === 'insert-row' || command.type === 'insert-column') {
    return index < command.index ? index : index + count;
  }
  if (index < command.index) return index;
  if (index >= command.index + count) return index - count;
  return null;
}

function transformRange(
  range: CellRange,
  axis: 'row' | 'column',
  command: Extract<
    WorkbookCommand,
    { readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' }
  >,
): CellRange | null {
  const start = range.start[axis];
  const end = range.end[axis];
  const count = command.count ?? 1;
  let nextStart: number;
  let nextEnd: number;
  if (command.type === 'insert-row' || command.type === 'insert-column') {
    nextStart = start < command.index ? start : start + count;
    nextEnd = end < command.index ? end : end + count;
  } else {
    const deletionEnd = command.index + count - 1;
    if (end < command.index) {
      nextStart = start;
      nextEnd = end;
    } else if (start > deletionEnd) {
      nextStart = start - count;
      nextEnd = end - count;
    } else {
      nextStart = start < command.index ? start : command.index;
      nextEnd = end > deletionEnd ? end - count : command.index - 1;
      if (nextStart > nextEnd) return null;
    }
  }
  return {
    start: { ...range.start, [axis]: nextStart },
    end: { ...range.end, [axis]: nextEnd },
  };
}

function transformStructure(
  input: SpreadsheetDocumentInput,
  command: Extract<
    WorkbookCommand,
    { readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column' }
  >,
  sheetIds: readonly SheetId[],
  authoritativeValidations: Map<string, Map<string, ValidationId | null>>,
): void {
  const index = sheetIndex(sheetIds, command.sheet);
  const sheet = input.workbook.sheets[index];
  if (sheet === undefined) return;
  const axis = command.type.endsWith('row') ? 'row' : 'column';
  sheet.cells = sheet.cells.flatMap((item) => {
    const next = transformIndex(item[axis], command);
    return next === null ? [] : [{ ...item, [axis]: next }];
  });
  authoritativeValidations.set(
    sheet.id,
    new Map(
      sheet.cells
        .filter((item) => item.cell.validationId !== undefined)
        .map((item) => [cellKey(item.row, item.column), item.cell.validationId as ValidationId]),
    ),
  );
  if (axis === 'row') {
    sheet.rows = (sheet.rows ?? []).flatMap((row) => {
      const next = transformIndex(row.index, command);
      return next === null ? [] : [{ ...row, index: next }];
    });
  } else {
    sheet.columns = (sheet.columns ?? []).flatMap((column) => {
      const next = transformIndex(column.index, command);
      return next === null ? [] : [{ ...column, index: next }];
    });
  }
  for (const template of input.templates) {
    if (template.sheetId !== sheet.id || template.range === undefined) continue;
    const transformed = transformRange(template.range, axis, command);
    if (transformed === null) delete template.range;
    else {
      template.range = {
        sheetId: template.range.sheetId,
        ...transformed,
      };
    }
  }
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
  const range = internalPasteRange(
    source,
    command.target.range,
    command.type === 'paste-internal' && command.cut,
  );
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
      setCell(targetSheet, row, column, mapPasteCell(targetCell, sourceCell, command.mode));
      if (sourceCell?.input.type === 'custom' && command.mode !== 'format') {
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

export function prepareSchemaCommand(
  document: SpreadsheetDocument,
  command: WorkbookCommand,
  sheetIds: readonly SheetId[],
): SchemaCommandPlan {
  const input = structuredClone(document) as unknown as SpreadsheetDocumentInput;
  const authoritativeInputs = new Map<string, Set<string>>();
  const authoritativeValidations = new Map<string, Map<string, ValidationId | null>>();
  switch (command.type) {
    case 'insert-row':
    case 'delete-row':
    case 'insert-column':
    case 'delete-column':
      transformStructure(input, command, sheetIds, authoritativeValidations);
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
  }
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) {
    throw new TypeError(
      `Schema command plan produced an invalid spreadsheet document: ${JSON.stringify(parsed.diagnostics)}`,
    );
  }
  return { document: parsed.document, authoritativeInputs, authoritativeValidations };
}
