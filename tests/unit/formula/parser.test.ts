import { describe, expect, it } from 'vitest';
import { parseFormula } from '../../../src/formula/parser';
import { resolveFormulaReferences } from '../../../src/formula/reference-resolver';
import { renderFormula, translateFormula } from '../../../src/formula/reference-transform';
import { formulaDocument } from './helpers';

describe('typed formula parser', () => {
  it('preserves source spans and relative, absolute, mixed, range, and cross-sheet references', () => {
    const source = "='Data Sheet'!$A1+B$2+$C$3+D4:E5";
    const ast = parseFormula(source);
    expect(ast.kind).toBe('binary');
    expect(ast.span).toEqual({ start: 1, end: source.length });
    expect(JSON.stringify(ast)).toContain('"sheetToken":"Data Sheet"');
    expect(JSON.stringify(ast)).toContain('"columnAbsolute":true');
    expect(JSON.stringify(ast)).toContain('"rowAbsolute":true');
    expect(JSON.stringify(ast)).toContain('"kind":"range"');
  });

  it('resolves display sheet tokens to stable sheet ids and reports invalid references', () => {
    const document = formulaDocument([
      { id: 'main-id', name: 'Main', cells: [] },
      { id: 'data-id', name: 'Data Sheet', cells: [] },
    ]);
    const resolved = resolveFormulaReferences(
      parseFormula("='Data Sheet'!A1+Missing!A1"),
      document,
      'main-id',
    );
    expect(resolved.diagnostics).toEqual([
      expect.objectContaining({ code: 'FORMULA_REFERENCE_INVALID' }),
    ]);
    expect(JSON.stringify(resolved.ast)).toContain('"sheetId":"data-id"');
  });

  it('rejects malformed formulas with a stable parse diagnostic', () => {
    expect(() => parseFormula('=SUM(A1,')).toThrow(
      expect.objectContaining({ code: 'FORMULA_PARSE_ERROR' }),
    );
  });

  it('does not misclassify function names containing digits as A1 references', () => {
    expect(parseFormula('=LOG10(100)').kind).toBe('call');
    expect(parseFormula('=DAYS360(A1,B1)')).toMatchObject({
      kind: 'call',
      name: 'DAYS360',
    });
  });

  it('distinguishes unquoted sheet names ending in digits from cell references', () => {
    expect(parseFormula('=Sheet1!A1')).toMatchObject({
      kind: 'reference',
      reference: { sheetToken: 'Sheet1', row: 0, column: 0 },
    });
  });

  it('rejects formulas beyond the fixed parser safety budget', () => {
    expect(() => parseFormula(`=${'1+'.repeat(5000)}1`)).toThrow(
      expect.objectContaining({ code: 'FORMULA_PARSE_ERROR' }),
    );
  });

  it('translates only relative axes during copy and normalizes source from the typed AST', () => {
    const translated = translateFormula(parseFormula('=$A1+B$1+$C$1'), {
      rowDelta: 2,
      columnDelta: 1,
    });
    expect(renderFormula(translated)).toBe('=$A3+C$1+$C$1');
  });

  it('returns deeply frozen AST snapshots', () => {
    const ast = parseFormula('=SUM(A1,2)');
    expect(Object.isFrozen(ast)).toBe(true);
    expect(ast.kind === 'call' && Object.isFrozen(ast.arguments)).toBe(true);
    expect(ast.kind === 'call' && Object.isFrozen(ast.arguments[0])).toBe(true);
  });
});
