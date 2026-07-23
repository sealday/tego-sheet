import type {
  CellPoint,
  SheetInput,
  SheetRange,
  SpreadsheetDocumentInput,
} from '../../document/model/document';
import { parseA1Reference, renderA1Reference } from './a1';

export type CoordinateAxis = 'row' | 'column';
export type CoordinateTransformKind = 'insert' | 'delete';

interface FormulaTransformContext {
  readonly targetSheetName: string;
  readonly transformUnqualified: boolean;
}

const FORMULA_REFERENCE =
  /^((?:'(?:[^']|'')+'|[A-Z_][A-Z0-9_.]*)!)?(\$?[A-Z]+\$?[1-9]\d*)(?::((?:'(?:[^']|'')+'|[A-Z_][A-Z0-9_.]*)!)?(\$?[A-Z]+\$?[1-9]\d*))?/i;

function coordinate(point: CellPoint, axis: CoordinateAxis): number {
  return point[axis];
}

function withCoordinate(point: CellPoint, axis: CoordinateAxis, value: number): CellPoint {
  return axis === 'row' ? { row: value, column: point.column } : { row: point.row, column: value };
}

function assertTransformInput(axis: CoordinateAxis, index: number, count: number): void {
  if (axis !== 'row' && axis !== 'column') throw new TypeError('Coordinate axis is invalid');
  if (!Number.isSafeInteger(index) || index < 0) {
    throw new RangeError('Coordinate transform index must be a non-negative safe integer');
  }
  if (!Number.isSafeInteger(count) || count < 1) {
    throw new RangeError('Coordinate transform count must be a positive safe integer');
  }
  if (!Number.isSafeInteger(index + count)) {
    throw new RangeError('Coordinate transform extent must be a safe integer');
  }
}

/**
 * One immutable row/column insertion or deletion mapping shared by every document coordinate.
 */
export class CoordinateTransform {
  private constructor(
    readonly kind: CoordinateTransformKind,
    readonly axis: CoordinateAxis,
    readonly index: number,
    readonly count: number,
  ) {}

  static insert(axis: CoordinateAxis, index: number, count = 1): CoordinateTransform {
    assertTransformInput(axis, index, count);
    return new CoordinateTransform('insert', axis, index, count);
  }

  static delete(axis: CoordinateAxis, index: number, count = 1): CoordinateTransform {
    assertTransformInput(axis, index, count);
    return new CoordinateTransform('delete', axis, index, count);
  }

  point(point: CellPoint): CellPoint | null {
    const next = this.scalar(coordinate(point, this.axis));
    return next === null ? null : withCoordinate(point, this.axis, next);
  }

  scalar(value: number): number | null {
    if (this.kind === 'insert') return value < this.index ? value : value + this.count;
    if (value < this.index) return value;
    if (value >= this.index + this.count) return value - this.count;
    return null;
  }

  boundary(value: number): number {
    return this.scalar(value) ?? this.index;
  }

  range(range: SheetRange): SheetRange | null {
    const start = coordinate(range.start, this.axis);
    const end = coordinate(range.end, this.axis);
    if (this.kind === 'insert') {
      if (this.index <= start) {
        return {
          start: withCoordinate(range.start, this.axis, start + this.count),
          end: withCoordinate(range.end, this.axis, end + this.count),
        };
      }
      if (this.index <= end) {
        return {
          start: range.start,
          end: withCoordinate(range.end, this.axis, end + this.count),
        };
      }
      return range;
    }
    const deletionEnd = this.index + this.count - 1;
    if (end < this.index) return range;
    if (start > deletionEnd) {
      return {
        start: withCoordinate(range.start, this.axis, start - this.count),
        end: withCoordinate(range.end, this.axis, end - this.count),
      };
    }
    const keepsBefore = start < this.index;
    const keepsAfter = end > deletionEnd;
    if (!keepsBefore && !keepsAfter) return null;
    const nextStart = keepsBefore ? start : this.index;
    const nextEnd = keepsAfter ? end - this.count : this.index - 1;
    if (nextEnd < nextStart) return null;
    return {
      start: withCoordinate(range.start, this.axis, nextStart),
      end: withCoordinate(range.end, this.axis, nextEnd),
    };
  }

  formula(
    source: string,
    context: FormulaTransformContext = {
      targetSheetName: '',
      transformUnqualified: true,
    },
  ): string {
    if (!source.startsWith('=')) return source;
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
      if (!inString) {
        const match = FORMULA_REFERENCE.exec(source.slice(index));
        const before = source[index - 1];
        if (
          match !== null &&
          (before === undefined || !/[A-Z0-9_$]/i.test(before)) &&
          !/[A-Z0-9_$]/i.test(source[index + match[0].length] ?? '')
        ) {
          const firstQualifier = match[1];
          const secondQualifier = match[3];
          output += `${firstQualifier ?? ''}${this.transformFormulaReference(
            match[2] as string,
            firstQualifier,
            context,
          )}`;
          if (match[4] !== undefined) {
            output += `:${secondQualifier ?? ''}${this.transformFormulaReference(
              match[4],
              secondQualifier ?? firstQualifier,
              context,
            )}`;
          }
          index += match[0].length;
          continue;
        }
      }
      output += character;
      index += 1;
    }
    return output;
  }

  private transformFormulaReference(
    token: string,
    qualifier: string | undefined,
    context: FormulaTransformContext,
  ): string {
    if (
      qualifier === undefined
        ? !context.transformUnqualified
        : normalizeSheetQualifier(qualifier) !== context.targetSheetName.toLowerCase()
    ) {
      return token;
    }
    const reference = parseA1Reference(token.toUpperCase());
    const current = this.axis === 'row' ? reference.row : reference.column;
    const next = this.scalar(current);
    if (next === null) return '#REF!';
    if (next === current) return token;
    return renderA1Reference(
      this.axis === 'row' ? { ...reference, row: next } : { ...reference, column: next },
    );
  }
}

function normalizeSheetQualifier(qualifier: string): string {
  const name = qualifier.slice(0, -1);
  return (name.startsWith("'") ? name.slice(1, -1).replaceAll("''", "'") : name).toLowerCase();
}

function transformCellFormula(
  cell: SheetInput['cells'][number]['cell'],
  transform: CoordinateTransform,
  context: FormulaTransformContext,
): SheetInput['cells'][number]['cell'] {
  if (cell.input.type !== 'formula') return cell;
  return {
    ...cell,
    input: {
      ...cell.input,
      source: transform.formula(cell.input.source, context),
    },
  };
}

/** Applies a coordinate mapping to every coordinate-bearing field currently in a schema 2 sheet. */
export function transformSheetCoordinates(
  sheet: SheetInput,
  transform: CoordinateTransform,
  context: FormulaTransformContext = {
    targetSheetName: sheet.name,
    transformUnqualified: true,
  },
): SheetInput {
  if (
    transform.kind === 'delete' &&
    ((transform.axis === 'row' &&
      sheet.rowCount !== undefined &&
      transform.index + transform.count > sheet.rowCount) ||
      (transform.axis === 'column' &&
        sheet.columnCount !== undefined &&
        transform.index + transform.count > sheet.columnCount))
  ) {
    throw new RangeError('Coordinate deletion exceeds the declared sheet structure');
  }
  if (
    transform.kind === 'insert' &&
    ((transform.axis === 'row' &&
      sheet.rowCount !== undefined &&
      !Number.isSafeInteger(sheet.rowCount + transform.count)) ||
      (transform.axis === 'column' &&
        sheet.columnCount !== undefined &&
        !Number.isSafeInteger(sheet.columnCount + transform.count)))
  ) {
    throw new RangeError('Coordinate insertion exceeds the safe sheet structure');
  }
  const cells = sheet.cells.flatMap((item) => {
    const point = transform.point({ row: item.row, column: item.column });
    return point === null
      ? []
      : [
          {
            ...item,
            ...point,
            cell: transformCellFormula(item.cell, transform, context),
          },
        ];
  });
  const merges = sheet.merges.flatMap((range) => {
    const next = transform.range(range);
    return next === null ? [] : [next];
  });
  const freeze =
    sheet.freeze === undefined
      ? undefined
      : transform.axis === 'row'
        ? { ...sheet.freeze, row: transform.boundary(sheet.freeze.row) }
        : { ...sheet.freeze, column: transform.boundary(sheet.freeze.column) };
  const filterRange = sheet.filter?.range && transform.range(sheet.filter.range);
  const filter =
    sheet.filter === undefined
      ? undefined
      : {
          ...sheet.filter,
          ...(filterRange === undefined || filterRange === null ? {} : { range: filterRange }),
          ...(filterRange === null ? { range: undefined } : {}),
          ...(transform.axis === 'column'
            ? {
                filters: sheet.filter.filters.flatMap((item) => {
                  const column = transform.scalar(item.column);
                  return column === null ? [] : [{ ...item, column }];
                }),
                sort:
                  sheet.filter.sort == null
                    ? sheet.filter.sort
                    : (() => {
                        const column = transform.scalar(sheet.filter!.sort!.column);
                        return column === null ? null : { ...sheet.filter!.sort!, column };
                      })(),
              }
            : {}),
        };
  return {
    ...sheet,
    ...(transform.axis === 'row' && sheet.rowCount !== undefined
      ? {
          rowCount:
            transform.kind === 'insert'
              ? sheet.rowCount + transform.count
              : sheet.rowCount - transform.count,
        }
      : {}),
    ...(transform.axis === 'column' && sheet.columnCount !== undefined
      ? {
          columnCount:
            transform.kind === 'insert'
              ? sheet.columnCount + transform.count
              : sheet.columnCount - transform.count,
        }
      : {}),
    cells,
    merges,
    ...(transform.axis === 'row'
      ? {
          rows: (sheet.rows ?? []).flatMap((item) => {
            const index = transform.scalar(item.index);
            return index === null ? [] : [{ ...item, index }];
          }),
        }
      : {
          columns: (sheet.columns ?? []).flatMap((item) => {
            const index = transform.scalar(item.index);
            return index === null ? [] : [{ ...item, index }];
          }),
        }),
    ...(freeze === undefined ? {} : { freeze }),
    ...(filter === undefined ? {} : { filter }),
  };
}

function transformFormulaCells(
  sheet: SheetInput,
  transform: CoordinateTransform,
  context: FormulaTransformContext,
): SheetInput {
  return {
    ...sheet,
    cells: sheet.cells.map((item) => ({
      ...item,
      cell: transformCellFormula(item.cell, transform, context),
    })),
  };
}

/**
 * Applies one structural coordinate mapping across the target sheet, cross-sheet formulas,
 * and document-level template/print ranges.
 */
export function transformDocumentCoordinates(
  document: SpreadsheetDocumentInput,
  targetSheetId: string,
  transform: CoordinateTransform,
): SpreadsheetDocumentInput {
  const output = structuredClone(document);
  const target = output.workbook.sheets.find((sheet) => sheet.id === targetSheetId);
  if (target === undefined) return output;
  output.workbook.sheets = output.workbook.sheets.map((sheet) =>
    sheet.id === targetSheetId
      ? transformSheetCoordinates(sheet, transform, {
          targetSheetName: target.name,
          transformUnqualified: true,
        })
      : transformFormulaCells(sheet, transform, {
          targetSheetName: target.name,
          transformUnqualified: false,
        }),
  );
  output.templates = output.templates.map((template) => {
    if (template.sheetId !== targetSheetId || template.range === undefined) return template;
    const range = transform.range(template.range);
    if (range === null) {
      const { range: _range, ...withoutRange } = template;
      return withoutRange;
    }
    return {
      ...template,
      range: {
        sheetId: template.range.sheetId,
        ...range,
      },
    };
  });
  return output;
}
