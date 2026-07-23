import type { SpreadsheetDocument } from '../document';
import type { FormulaAst, FormulaDiagnostic, FormulaReference } from './ast';
import { freezeFormulaAst } from './ast';

/** Resolved AST and any stable-reference diagnostics. */
export interface ReferenceResolutionResult {
  /** AST with stable sheet identifiers attached to valid references. */
  readonly ast: FormulaAst;
  /** Unknown-sheet and invalid-reference diagnostics. */
  readonly diagnostics: readonly FormulaDiagnostic[];
}

function resolveReference(
  reference: FormulaReference,
  document: SpreadsheetDocument,
  currentSheetId: string,
  diagnostics: FormulaDiagnostic[],
): FormulaReference {
  let sheet;
  if (reference.sheetToken === undefined) {
    sheet = document.workbook.sheets.find(({ id }) => id === currentSheetId);
  } else {
    const exact = document.workbook.sheets.find(({ name }) => name === reference.sheetToken);
    const normalized = document.workbook.sheets.filter(
      ({ name }) => name.toLowerCase() === reference.sheetToken?.toLowerCase(),
    );
    sheet = exact ?? (normalized.length === 1 ? normalized[0] : undefined);
    if (exact === undefined && normalized.length > 1) {
      diagnostics.push({
        code: 'FORMULA_REFERENCE_INVALID',
        message: `Ambiguous sheet ${reference.sheetToken}`,
      });
      return reference;
    }
  }
  if (sheet === undefined) {
    diagnostics.push({
      code: 'FORMULA_REFERENCE_INVALID',
      message: `Unknown sheet ${reference.sheetToken ?? currentSheetId}`,
    });
    return reference;
  }
  return { ...reference, sheetId: sheet.id };
}

function visit(
  ast: FormulaAst,
  document: SpreadsheetDocument,
  currentSheetId: string,
  diagnostics: FormulaDiagnostic[],
): FormulaAst {
  if (ast.kind === 'reference') {
    return {
      ...ast,
      reference: resolveReference(ast.reference, document, currentSheetId, diagnostics),
    };
  }
  if (ast.kind === 'range') {
    return {
      ...ast,
      start: resolveReference(ast.start, document, currentSheetId, diagnostics),
      end: resolveReference(ast.end, document, currentSheetId, diagnostics),
    };
  }
  if (ast.kind === 'binary') {
    return {
      ...ast,
      left: visit(ast.left, document, currentSheetId, diagnostics),
      right: visit(ast.right, document, currentSheetId, diagnostics),
    };
  }
  if (ast.kind === 'unary') {
    return { ...ast, operand: visit(ast.operand, document, currentSheetId, diagnostics) };
  }
  if (ast.kind === 'call') {
    return {
      ...ast,
      arguments: ast.arguments.map((argument) =>
        visit(argument, document, currentSheetId, diagnostics),
      ),
    };
  }
  return ast;
}

/** Resolves sheet display tokens against stable Workbook 2.0 sheet identifiers. */
export function resolveFormulaReferences(
  ast: FormulaAst,
  document: SpreadsheetDocument,
  currentSheetId: string,
): ReferenceResolutionResult {
  const diagnostics: FormulaDiagnostic[] = [];
  return {
    ast: freezeFormulaAst(visit(ast, document, currentSheetId, diagnostics)),
    diagnostics: Object.freeze(diagnostics),
  };
}
