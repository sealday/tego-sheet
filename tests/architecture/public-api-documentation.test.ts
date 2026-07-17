import { join } from 'node:path';
import ts from 'typescript';
import { Application, normalizePath, TSConfigReader } from 'typedoc';
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

const structuredAliases = {
  AutoFilterData: 'ref',
  AutoFilterItemData: 'ci',
  AutoFilterSortData: 'ci',
  CellBorders: 'top',
  CellData: 'text',
  CellStyle: 'format',
  CellsData: null,
  ColsData: 'len',
  ColumnData: 'width',
  FontStyle: 'name',
  RowData: 'cells',
  RowsData: 'len',
  SheetData: 'rows',
  ValidationData: 'refs',
} as const;

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

const isRecord = (value: unknown): value is Readonly<Record<string, unknown>> =>
  typeof value === 'object' && value !== null;

const findNested = (
  value: unknown,
  predicate: (record: Readonly<Record<string, unknown>>) => boolean,
): Readonly<Record<string, unknown>> | undefined => {
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findNested(item, predicate);
      if (found !== undefined) return found;
    }
    return undefined;
  }
  if (!isRecord(value)) return undefined;
  if (predicate(value)) return value;
  for (const item of Object.values(value)) {
    const found = findNested(item, predicate);
    if (found !== undefined) return found;
  }
  return undefined;
};

const summaryText = (reflection: Readonly<Record<string, unknown>>): string => {
  const comment = reflection.comment;
  if (!isRecord(comment) || !Array.isArray(comment.summary)) return '';
  return comment.summary
    .map((part) => (isRecord(part) && typeof part.text === 'string' ? part.text : ''))
    .join('')
    .trim();
};

const hasSummary = (reflection: Readonly<Record<string, unknown>>): boolean =>
  summaryText(reflection) !== '';

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

  it('converts structured aliases into documented TypeDoc members instead of self references', async () => {
    const app = await Application.bootstrap(
      {
        entryPoints: ['src/index.ts'],
        excludeInternal: true,
        excludePrivate: true,
        excludeProtected: true,
        readme: 'none',
        skipErrorChecking: true,
        tsconfig: 'tsconfig.json',
      },
      [new TSConfigReader()],
    );
    const project = await app.convert();
    if (project === undefined) throw new Error('TypeDoc must convert the public entry point');
    const serialized = app.serializer.projectToObject(project, normalizePath(root));
    const aliases = new Map(
      (serialized.children ?? [])
        .filter((child) => Object.hasOwn(structuredAliases, child.name))
        .map((child) => [child.name, child]),
    );

    expect([...aliases.keys()].sort()).toEqual(Object.keys(structuredAliases).sort());

    const selfReferences: string[] = [];
    const missingDocumentedMembers: string[] = [];
    for (const [name, member] of Object.entries(structuredAliases)) {
      const reflection = aliases.get(name);
      if (reflection?.type?.type === 'reference' && reflection.type.name === name) {
        selfReferences.push(name);
      }
      if (member === null) {
        const indexSignature = findNested(
          reflection,
          (value) => Array.isArray(value.indexSignatures) && value.indexSignatures.length > 0,
        );
        if (indexSignature === undefined) missingDocumentedMembers.push(`${name}.[key]`);
        continue;
      }
      const documentedMember = findNested(
        reflection,
        (value) => value.variant === 'declaration' && value.name === member && hasSummary(value),
      );
      if (documentedMember === undefined) missingDocumentedMembers.push(`${name}.${member}`);
    }

    expect(selfReferences).toEqual([]);
    expect(missingDocumentedMembers).toEqual([]);

    const activeSheetChange = (serialized.children ?? []).find(
      (child) => child.name === 'ActiveSheetChangeEvent',
    );
    const activeSheet = findNested(
      activeSheetChange,
      (value) => value.variant === 'declaration' && value.name === 'sheet',
    );
    expect(activeSheet).toBeDefined();
    expect(summaryText(activeSheet ?? {})).toBe(
      'Worksheet reported as active by the activation event.',
    );
  }, 15_000);
});
