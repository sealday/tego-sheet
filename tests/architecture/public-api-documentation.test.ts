import { join } from 'node:path';
import ts from 'typescript';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const entryPoint = join(root, 'src/index.ts');

const expectedPublicExports = [
  'ActiveSheetChangeEvent',
  'AutoFilterData',
  'AutoFilterItemData',
  'AutoFilterSortData',
  'BorderLine',
  'BorderMode',
  'CellAddress',
  'CellBorders',
  'CellData',
  'CellEditEvent',
  'CellPoint',
  'CellRange',
  'CellStyle',
  'CellsData',
  'ChangeSource',
  'ColsData',
  'ColumnData',
  'FilterDefinition',
  'FontStyle',
  'HorizontalAlign',
  'JsonValue',
  'LocaleDefinition',
  'LocaleMessages',
  'PasteEvent',
  'RowData',
  'RowsData',
  'Selection',
  'SheetColumnOptions',
  'SheetData',
  'SheetId',
  'SheetOptions',
  'SheetRowOptions',
  'SheetTabItem',
  'SheetTabsRenderProps',
  'SheetTabsRenderer',
  'TegoSheet',
  'TegoSheetError',
  'TegoSheetErrorCode',
  'TegoSheetException',
  'TegoSheetHandle',
  'TegoSheetProps',
  'ToolbarAction',
  'ToolbarRenderProps',
  'ToolbarRenderer',
  'ValidationData',
  'ValidationIssue',
  'ValidationOperator',
  'ValidationResult',
  'ValidationRule',
  'ValidationType',
  'VerticalAlign',
  'WorkbookChange',
  'WorkbookChangeKind',
  'WorkbookData',
  'WorkbookInput',
] as const;

interface PublicDeclaration {
  readonly name: string;
  readonly symbol?: ts.Symbol;
  readonly node?: ts.Node;
}

const createPublicProgram = (): ts.Program => {
  const configPath = ts.findConfigFile(root, ts.sys.fileExists, 'tsconfig.json');
  if (configPath === undefined) throw new Error('tsconfig.json must exist');
  const config = ts.readConfigFile(configPath, ts.sys.readFile);
  if (config.error !== undefined) {
    throw new Error(ts.flattenDiagnosticMessageText(config.error.messageText, '\n'));
  }
  const parsed = ts.parseJsonConfigFileContent(config.config, ts.sys, root);
  return ts.createProgram({ rootNames: parsed.fileNames, options: parsed.options });
};

const hasDocumentation = (declaration: PublicDeclaration, checker: ts.TypeChecker): boolean => {
  if (declaration.symbol !== undefined) {
    return (
      ts.displayPartsToString(declaration.symbol.getDocumentationComment(checker)).trim() !== ''
    );
  }
  return declaration.node !== undefined && ts.getJSDocCommentsAndTags(declaration.node).length > 0;
};

const publicDeclarations = (
  exportedSymbol: ts.Symbol,
  checker: ts.TypeChecker,
): readonly PublicDeclaration[] => {
  const symbol =
    exportedSymbol.flags & ts.SymbolFlags.Alias
      ? checker.getAliasedSymbol(exportedSymbol)
      : exportedSymbol;
  const declarations: PublicDeclaration[] = [{ name: exportedSymbol.name, symbol }];
  const seen = new Set<ts.Symbol>([symbol]);
  const addSymbol = (member: ts.Symbol, name: string): void => {
    if (seen.has(member)) return;
    seen.add(member);
    declarations.push({ name, symbol: member });
  };

  if (symbol.flags & (ts.SymbolFlags.Class | ts.SymbolFlags.Interface)) {
    for (const member of checker.getDeclaredTypeOfSymbol(symbol).getProperties()) {
      if (
        !member.declarations?.some((declaration) =>
          declaration.getSourceFile().fileName.startsWith(join(root, 'src')),
        )
      ) {
        continue;
      }
      addSymbol(member, `${exportedSymbol.name}.${member.name}`);
    }
  }

  for (const declaration of symbol.declarations ?? []) {
    if (ts.isClassDeclaration(declaration)) {
      for (const member of declaration.members) {
        if (ts.isConstructorDeclaration(member)) {
          declarations.push({ name: `${exportedSymbol.name}.constructor`, node: member });
        }
      }
    }
    if (!ts.isTypeAliasDeclaration(declaration)) continue;
    const visit = (node: ts.Node): void => {
      if (ts.isTypeLiteralNode(node)) {
        for (const member of node.members) {
          if (member.name === undefined) continue;
          const memberSymbol = checker.getSymbolAtLocation(member.name);
          if (memberSymbol !== undefined) {
            addSymbol(memberSymbol, `${exportedSymbol.name}.${member.name.getText()}`);
          }
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(declaration.type);
  }

  return declarations;
};

describe('public API documentation', () => {
  it('documents the exact root export surface and its public members', () => {
    const program = createPublicProgram();
    const checker = program.getTypeChecker();
    const source = program.getSourceFile(entryPoint);
    if (source === undefined)
      throw new Error('src/index.ts must be part of the TypeScript program');
    const moduleSymbol = checker.getSymbolAtLocation(source);
    if (moduleSymbol === undefined) throw new Error('src/index.ts must be a module');
    const exports = checker
      .getExportsOfModule(moduleSymbol)
      .sort((left, right) => (left.name < right.name ? -1 : left.name > right.name ? 1 : 0));

    expect(exports.map((symbol) => symbol.name)).toEqual(expectedPublicExports);

    const undocumented = exports
      .flatMap((symbol) => publicDeclarations(symbol, checker))
      .filter((declaration) => !hasDocumentation(declaration, checker))
      .map((declaration) => declaration.name);

    expect(undocumented).toEqual([]);
  });
});
