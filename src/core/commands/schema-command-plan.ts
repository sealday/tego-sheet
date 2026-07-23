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

type ValidationId = NonNullable<Cell['validationId']>;

export interface SchemaCommandPlan {
  readonly document: SpreadsheetDocument;
  readonly authoritativeInputs: ReadonlyMap<string, ReadonlySet<string>>;
  readonly authoritativeValidations: ReadonlyMap<string, ReadonlyMap<string, ValidationId | null>>;
}

export interface SchemaProjectionCommit {
  readonly result: unknown;
  readonly kind: Extract<WorkbookChangeKind, 'clipboard' | 'autofill'>;
  readonly sheet: SheetId;
  readonly range: CellRange;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function sheetIndex(sheetIds: readonly SheetId[], sheet: SheetId): number {
  return sheetIds.findIndex((candidate) => candidate === sheet);
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
      if (command.type === 'paste-internal' && command.cut && sourceCell === undefined) continue;
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

export function prepareSchemaProjectionCommit(
  command: Extract<WorkbookCommand, { readonly type: 'paste-internal' | 'autofill' }>,
  projection: WorkbookData,
  sheetIds: readonly SheetId[],
  capturePasteValues: boolean,
): SchemaProjectionCommit {
  const range = internalPasteRange(
    command.source.range,
    command.target.range,
    command.type === 'paste-internal' && command.cut,
  );
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
  }
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) {
    throw new TypeError(
      `Schema command plan produced an invalid spreadsheet document: ${JSON.stringify(parsed.diagnostics)}`,
    );
  }
  return { document: parsed.document, authoritativeInputs, authoritativeValidations };
}
