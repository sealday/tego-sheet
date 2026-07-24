import type { Cell, CellInput, SheetRange } from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';

function positiveModulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor;
}

function parseIsoDate(value: string): number | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/u.exec(value);
  if (match === null) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const timestamp = Date.UTC(year, month - 1, day);
  const date = new Date(timestamp);
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
    ? timestamp
    : undefined;
}

function inferredInput(seeds: readonly CellInput[], ordinal: number): CellInput | undefined {
  if (seeds.length < 2) return undefined;
  if (seeds.every((input) => input.type === 'number')) {
    const [first, second] = seeds as readonly [
      Extract<CellInput, { readonly type: 'number' }>,
      Extract<CellInput, { readonly type: 'number' }>,
    ];
    const value = first.value + (second.value - first.value) * ordinal;
    return Number.isFinite(value) ? { type: 'number', value } : undefined;
  }
  if (!seeds.every((input) => input.type === 'string')) return undefined;
  const strings = seeds as readonly Extract<CellInput, { readonly type: 'string' }>[];
  const firstDate = parseIsoDate(strings[0]!.value);
  const secondDate = parseIsoDate(strings[1]!.value);
  if (firstDate !== undefined && secondDate !== undefined) {
    const timestamp = firstDate + (secondDate - firstDate) * ordinal;
    return { type: 'string', value: new Date(timestamp).toISOString().slice(0, 10) };
  }
  const matches = strings.map(({ value }) => /^(.*?)(-?\d+)$/u.exec(value));
  if (matches.some((match) => match === null)) return undefined;
  const first = matches[0]!;
  const second = matches[1]!;
  if (matches.some((match) => match![1] !== first[1])) return undefined;
  const start = Number(first[2]);
  const step = Number(second[2]) - start;
  const value = start + step * ordinal;
  const width = Math.max(...matches.map((match) => match![2]!.replace(/^-/, '').length));
  const digits = String(Math.abs(value)).padStart(width, '0');
  return {
    type: 'string',
    value: `${first[1]}${value < 0 ? '-' : ''}${digits}`,
  };
}

/** Creates one typed-pattern and AST formula resolver for an autofill operation. */
export function createTypedAutofillResolver(
  source: SheetRange,
  target: SheetRange,
  sourceCellAt: (row: number, column: number) => Cell | undefined,
): (row: number, column: number) => CellInput | undefined {
  const sourceRows = source.end.row - source.start.row + 1;
  const sourceColumns = source.end.column - source.start.column + 1;
  const vertical = target.start.row > source.end.row || target.end.row < source.start.row;
  const horizontal =
    target.start.column > source.end.column || target.end.column < source.start.column;
  const verticalSeeds = new Map<number, readonly CellInput[] | undefined>();
  const horizontalSeeds = new Map<number, readonly CellInput[] | undefined>();
  return (row, column) => {
    const sourceRow = source.start.row + positiveModulo(row - target.start.row, sourceRows);
    const sourceColumn =
      source.start.column + positiveModulo(column - target.start.column, sourceColumns);
    const sourceCell = sourceCellAt(sourceRow, sourceColumn);
    if (sourceCell?.input.type === 'formula') {
      try {
        return {
          type: 'formula',
          source: renderFormula(
            translateFormula(parseFormula(sourceCell.input.source), {
              rowDelta: row - sourceRow,
              columnDelta: column - sourceColumn,
            }),
          ),
        };
      } catch {
        return { type: 'formula', source: '=#REF!' };
      }
    }
    if (vertical) {
      if (!verticalSeeds.has(sourceColumn)) {
        const cells = Array.from({ length: sourceRows }, (_, index) =>
          sourceCellAt(source.start.row + index, sourceColumn),
        );
        verticalSeeds.set(
          sourceColumn,
          cells.every((cell) => cell !== undefined) ? cells.map((cell) => cell!.input) : undefined,
        );
      }
      const seeds = verticalSeeds.get(sourceColumn);
      const inferred =
        seeds === undefined ? undefined : inferredInput(seeds, row - source.start.row);
      if (inferred !== undefined) return inferred;
    } else if (horizontal) {
      if (!horizontalSeeds.has(sourceRow)) {
        const cells = Array.from({ length: sourceColumns }, (_, index) =>
          sourceCellAt(sourceRow, source.start.column + index),
        );
        horizontalSeeds.set(
          sourceRow,
          cells.every((cell) => cell !== undefined) ? cells.map((cell) => cell!.input) : undefined,
        );
      }
      const seeds = horizontalSeeds.get(sourceRow);
      const inferred =
        seeds === undefined ? undefined : inferredInput(seeds, column - source.start.column);
      if (inferred !== undefined) return inferred;
    }
    return sourceCell?.input === undefined ? undefined : structuredClone(sourceCell.input);
  };
}
