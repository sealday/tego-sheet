import type { FormulaAst, FormulaReference } from './ast';
import { freezeFormulaAst } from './ast';

/** Relative row and column movement applied during copy or fill. */
export interface FormulaTranslation {
  /** Signed row displacement. */
  readonly rowDelta: number;
  /** Signed column displacement. */
  readonly columnDelta: number;
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
  if (ast.kind === 'reference') return renderReference(ast.reference);
  if (ast.kind === 'range') return `${renderReference(ast.start)}:${renderReference(ast.end)}`;
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
