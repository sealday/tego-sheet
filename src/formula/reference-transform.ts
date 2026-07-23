import type { FormulaAst, FormulaReference } from './ast';
import { freezeFormulaAst } from './ast';

/** Relative row and column movement applied during copy or fill. */
export interface FormulaTranslation {
  /** Signed row displacement. */
  readonly rowDelta: number;
  /** Signed column displacement. */
  readonly columnDelta: number;
}

/** Minimal coordinate mapper consumed by typed structural formula transforms. */
export interface FormulaCoordinateMapper {
  /** Maps one cell or returns null when the cell was deleted. */
  readonly transformPoint: (point: {
    readonly row: number;
    readonly column: number;
  }) => { readonly row: number; readonly column: number } | null;
  /** Maps one inclusive range or returns null when the complete range was deleted. */
  readonly transformRange: (range: {
    readonly start: { readonly row: number; readonly column: number };
    readonly end: { readonly row: number; readonly column: number };
  }) => {
    readonly start: { readonly row: number; readonly column: number };
    readonly end: { readonly row: number; readonly column: number };
  } | null;
}

/** Sheet-selection rules for a structural coordinate transform. */
export interface FormulaCoordinateTransformContext {
  /** Display name of the structurally changed sheet. */
  readonly targetSheetName: string;
  /** Whether unqualified references address the changed sheet. */
  readonly transformUnqualified: boolean;
}

function columnLabel(column: number): string {
  let value = column + 1;
  let output = '';
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + (value % 26)) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

function sheetQualifier(token: string | undefined): string {
  if (token === undefined) return '';
  return /^[A-Z_][A-Z0-9_.]*$/iu.test(token) ? `${token}!` : `'${token.replaceAll("'", "''")}'!`;
}

function renderReference(reference: FormulaReference): string {
  return `${sheetQualifier(reference.sheetToken)}${reference.columnAbsolute ? '$' : ''}${columnLabel(reference.column)}${reference.rowAbsolute ? '$' : ''}${reference.row + 1}`;
}

function renderRangeEnd(start: FormulaReference, end: FormulaReference): string {
  if (start.sheetToken === end.sheetToken) {
    return `${end.columnAbsolute ? '$' : ''}${columnLabel(end.column)}${end.rowAbsolute ? '$' : ''}${end.row + 1}`;
  }
  return renderReference(end);
}

const precedence: Readonly<Record<string, number | undefined>> = {
  '=': 1,
  '==': 1,
  '<>': 1,
  '!=': 1,
  '>': 1,
  '>=': 1,
  '<': 1,
  '<=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
};

function renderNode(ast: FormulaAst, parentPrecedence = 0): string {
  if (ast.kind === 'number') return String(ast.value);
  if (ast.kind === 'string') return `"${ast.value.replaceAll('"', '""')}"`;
  if (ast.kind === 'boolean') return ast.value ? 'TRUE' : 'FALSE';
  if (ast.kind === 'error') return ast.value;
  if (ast.kind === 'reference') return renderReference(ast.reference);
  if (ast.kind === 'range')
    return `${renderReference(ast.start)}:${renderRangeEnd(ast.start, ast.end)}`;
  if (ast.kind === 'call')
    return `${ast.name}(${ast.arguments.map((argument) => renderNode(argument)).join(',')})`;
  if (ast.kind === 'unary') return `-${renderNode(ast.operand, 5)}`;
  const level = precedence[ast.operator] ?? 0;
  const rendered = `${renderNode(ast.left, level)}${ast.operator}${renderNode(ast.right, level + 1)}`;
  return level < parentPrecedence ? `(${rendered})` : rendered;
}

/** Serializes a typed formula AST into canonical source text. */
export function renderFormula(ast: FormulaAst): string {
  return `=${renderNode(ast)}`;
}

function translateReference(
  reference: FormulaReference,
  translation: FormulaTranslation,
): FormulaReference {
  const row = reference.rowAbsolute ? reference.row : reference.row + translation.rowDelta;
  const column = reference.columnAbsolute
    ? reference.column
    : reference.column + translation.columnDelta;
  if (row < 0 || column < 0) throw new RangeError('Formula translation produced #REF!');
  return { ...reference, row, column };
}

/** Applies copy/fill translation to relative axes while preserving stable sheet identity. */
export function translateFormula(ast: FormulaAst, translation: FormulaTranslation): FormulaAst {
  if (
    !Number.isSafeInteger(translation.rowDelta) ||
    !Number.isSafeInteger(translation.columnDelta)
  ) {
    throw new RangeError('Formula translation deltas must be safe integers');
  }
  if (ast.kind === 'reference') {
    return freezeFormulaAst({ ...ast, reference: translateReference(ast.reference, translation) });
  }
  if (ast.kind === 'range') {
    return freezeFormulaAst({
      ...ast,
      start: translateReference(ast.start, translation),
      end: translateReference(ast.end, translation),
    });
  }
  if (ast.kind === 'unary') {
    return freezeFormulaAst({ ...ast, operand: translateFormula(ast.operand, translation) });
  }
  if (ast.kind === 'binary') {
    return freezeFormulaAst({
      ...ast,
      left: translateFormula(ast.left, translation),
      right: translateFormula(ast.right, translation),
    });
  }
  if (ast.kind === 'call') {
    return freezeFormulaAst({
      ...ast,
      arguments: ast.arguments.map((argument) => translateFormula(argument, translation)),
    });
  }
  return ast;
}

function targetsSheet(
  reference: FormulaReference,
  context: FormulaCoordinateTransformContext,
): boolean {
  return reference.sheetToken === undefined
    ? context.transformUnqualified
    : reference.sheetToken.toLowerCase() === context.targetSheetName.toLowerCase();
}

/** Rewrites reference nodes through one F2 structural coordinate mapping. */
export function transformFormulaCoordinates(
  ast: FormulaAst,
  mapper: FormulaCoordinateMapper,
  context: FormulaCoordinateTransformContext,
): FormulaAst {
  if (ast.kind === 'reference') {
    if (!targetsSheet(ast.reference, context)) return ast;
    const point = mapper.transformPoint(ast.reference);
    return point === null
      ? freezeFormulaAst({ kind: 'error', value: '#REF!', span: ast.span })
      : freezeFormulaAst({ ...ast, reference: { ...ast.reference, ...point } });
  }
  if (ast.kind === 'range') {
    const targetsStart = targetsSheet(ast.start, context);
    const targetsEnd = targetsSheet(ast.end, context);
    if (!targetsStart && !targetsEnd) return ast;
    if (targetsStart && targetsEnd) {
      const range = mapper.transformRange({ start: ast.start, end: ast.end });
      return range === null
        ? freezeFormulaAst({ kind: 'error', value: '#REF!', span: ast.span })
        : freezeFormulaAst({
            ...ast,
            start: { ...ast.start, ...range.start },
            end: { ...ast.end, ...range.end },
          });
    }
    const start = targetsStart ? mapper.transformPoint(ast.start) : ast.start;
    const end = targetsEnd ? mapper.transformPoint(ast.end) : ast.end;
    return start === null || end === null
      ? freezeFormulaAst({ kind: 'error', value: '#REF!', span: ast.span })
      : freezeFormulaAst({
          ...ast,
          start: { ...ast.start, ...start },
          end: { ...ast.end, ...end },
        });
  }
  if (ast.kind === 'unary') {
    return freezeFormulaAst({
      ...ast,
      operand: transformFormulaCoordinates(ast.operand, mapper, context),
    });
  }
  if (ast.kind === 'binary') {
    return freezeFormulaAst({
      ...ast,
      left: transformFormulaCoordinates(ast.left, mapper, context),
      right: transformFormulaCoordinates(ast.right, mapper, context),
    });
  }
  if (ast.kind === 'call') {
    return freezeFormulaAst({
      ...ast,
      arguments: ast.arguments.map((argument) =>
        transformFormulaCoordinates(argument, mapper, context),
      ),
    });
  }
  return ast;
}

function renameReferenceSheet(
  reference: FormulaReference,
  previousName: string,
  nextName: string,
): FormulaReference {
  return reference.sheetToken?.toLowerCase() === previousName.toLowerCase()
    ? { ...reference, sheetToken: nextName }
    : reference;
}

/** Rewrites explicit sheet display tokens after a stable sheet is renamed. */
export function renameFormulaSheet(
  ast: FormulaAst,
  previousName: string,
  nextName: string,
): FormulaAst {
  if (ast.kind === 'reference') {
    const reference = renameReferenceSheet(ast.reference, previousName, nextName);
    return reference === ast.reference ? ast : freezeFormulaAst({ ...ast, reference });
  }
  if (ast.kind === 'range') {
    const start = renameReferenceSheet(ast.start, previousName, nextName);
    const end = renameReferenceSheet(ast.end, previousName, nextName);
    return start === ast.start && end === ast.end ? ast : freezeFormulaAst({ ...ast, start, end });
  }
  if (ast.kind === 'unary') {
    return freezeFormulaAst({
      ...ast,
      operand: renameFormulaSheet(ast.operand, previousName, nextName),
    });
  }
  if (ast.kind === 'binary') {
    return freezeFormulaAst({
      ...ast,
      left: renameFormulaSheet(ast.left, previousName, nextName),
      right: renameFormulaSheet(ast.right, previousName, nextName),
    });
  }
  if (ast.kind === 'call') {
    return freezeFormulaAst({
      ...ast,
      arguments: ast.arguments.map((argument) =>
        renameFormulaSheet(argument, previousName, nextName),
      ),
    });
  }
  return ast;
}
