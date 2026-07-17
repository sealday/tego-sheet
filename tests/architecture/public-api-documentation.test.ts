import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import ts from 'typescript';
import { Application, normalizePath } from 'typedoc';
import type { TypeDocOptions } from 'typedoc';
import { describe, expect, it } from 'vitest';
import { publicApiProjectionPluginPath } from '../../website/plugins/strict-typedoc-generation';

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
  AutoFilterData: ['filters', 'ref', 'sort'],
  AutoFilterItemData: ['ci', 'operator', 'value'],
  AutoFilterSortData: ['ci', 'order'],
  CellBorders: ['bottom', 'left', 'right', 'top'],
  CellData: ['editable', 'merge', 'printable', 'style', 'text', 'value'],
  CellStyle: [
    'align',
    'bgcolor',
    'border',
    'color',
    'font',
    'format',
    'strike',
    'textwrap',
    'underline',
    'valign',
  ],
  CellsData: [],
  ColsData: ['len'],
  ColumnData: ['hide', 'style', 'width'],
  FontStyle: ['bold', 'italic', 'name', 'size'],
  RowData: ['cells', 'height', 'hide', 'style'],
  RowsData: ['len'],
  SheetData: ['autofilter', 'cols', 'freeze', 'merges', 'name', 'rows', 'styles', 'validations'],
  ValidationData: ['mode', 'operator', 'refs', 'required', 'type', 'value'],
} as const;

const tegoSheetPropNames = [
  'className',
  'defaultValue',
  'initialActiveSheetIndex',
  'locale',
  'onActiveSheetChange',
  'onCellEdit',
  'onChange',
  'onError',
  'onPaste',
  'onSelectionChange',
  'options',
  'readOnly',
  'sheetTabs',
  'style',
  'toolbar',
  'value',
] as const;

const callbackNames = [
  'onActiveSheetChange',
  'onCellEdit',
  'onChange',
  'onError',
  'onPaste',
  'onSelectionChange',
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

const visitRecords = (
  value: unknown,
  visitor: (record: Readonly<Record<string, unknown>>) => void,
): void => {
  if (Array.isArray(value)) {
    for (const item of value) visitRecords(item, visitor);
    return;
  }
  if (!isRecord(value)) return;
  visitor(value);
  for (const item of Object.values(value)) visitRecords(item, visitor);
};

const directRecords = (
  reflection: Readonly<Record<string, unknown>> | undefined,
  key: 'children' | 'indexSignatures',
): readonly Readonly<Record<string, unknown>>[] => {
  const values = reflection?.[key];
  return Array.isArray(values) ? values.filter(isRecord) : [];
};

const localDeclarationNames = (program: ts.Program): ReadonlySet<string> => {
  const names = new Set<string>();
  for (const source of program.getSourceFiles()) {
    if (!source.fileName.startsWith(join(root, 'src'))) continue;
    const visit = (node: ts.Node): void => {
      if (
        (ts.isClassDeclaration(node) ||
          ts.isEnumDeclaration(node) ||
          ts.isFunctionDeclaration(node) ||
          ts.isInterfaceDeclaration(node) ||
          ts.isTypeAliasDeclaration(node)) &&
        node.name !== undefined
      ) {
        names.add(node.name.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return names;
};

const referenceRoot = (name: string): string => {
  const packageImport = name.match(/^import\(["']tego-sheet["']\)\.([^.[<]+)/u)?.[1];
  if (packageImport !== undefined) return packageImport;
  return name.replace(/^tego-sheet[./:#]+/u, '').split(/[.[<]/u, 1)[0] ?? name;
};

const localReferenceViolations = (
  serialized: Readonly<Record<string, unknown>>,
  localNames: ReadonlySet<string>,
): readonly string[] => {
  const allowed = new Set<string>(expectedPublicExports);
  const rootNameById = new Map<number, string>();
  for (const child of directRecords(serialized, 'children')) {
    if (typeof child.name !== 'string') continue;
    visitRecords(child, (record) => {
      if (typeof record.id === 'number') rootNameById.set(record.id, child.name as string);
    });
  }

  const symbolIdMap = isRecord(serialized.symbolIdMap) ? serialized.symbolIdMap : {};
  const violations = new Set<string>();
  visitRecords(serialized, (record) => {
    if (record.type !== 'reference' || typeof record.name !== 'string') return;
    const target = record.target;
    let localRoot: string | undefined;

    if (typeof target === 'number' && target >= 0) {
      localRoot = rootNameById.get(target);
      if (localRoot === undefined) {
        const symbol = symbolIdMap[String(target)];
        if (isRecord(symbol) && symbol.packageName === 'tego-sheet') {
          const qualifiedName =
            typeof symbol.qualifiedName === 'string' ? symbol.qualifiedName : record.name;
          localRoot = referenceRoot(qualifiedName);
        }
      }
    } else if (isRecord(target) && target.packageName === 'tego-sheet') {
      const qualifiedName =
        typeof target.qualifiedName === 'string' ? target.qualifiedName : record.name;
      localRoot = referenceRoot(qualifiedName);
    } else {
      const nameRoot = referenceRoot(record.name);
      const packageName = typeof record.package === 'string' ? record.package : undefined;
      const hasLocalPrefix =
        record.name.startsWith('tego-sheet') || record.name.includes('import("tego-sheet")');
      if (
        packageName === 'tego-sheet' ||
        hasLocalPrefix ||
        (target === -1 && localNames.has(nameRoot))
      ) {
        localRoot = nameRoot;
      }
    }

    if (localRoot !== undefined && !allowed.has(localRoot)) {
      violations.add(`${record.name} -> ${localRoot}`);
    }
  });
  return [...violations].sort();
};

const findFile = async (directory: string, fileName: string): Promise<string | undefined> => {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) {
      const nested = await findFile(path, fileName);
      if (nested !== undefined) return nested;
    } else if (basename(path) === fileName) {
      return path;
    }
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

  it('@interface is TypeDoc display-only and keeps all 14 compiler declarations as type aliases', () => {
    const program = createPublicProgram();
    const source = program.getSourceFile(join(root, 'src/core/types/workbook.ts'));
    if (source === undefined)
      throw new Error('workbook types must be part of the TypeScript program');
    const declarations = new Map(
      source.statements
        .filter(ts.isTypeAliasDeclaration)
        .map((declaration) => [declaration.name.text, declaration]),
    );

    expect(
      [...declarations.keys()].filter((name) => Object.hasOwn(structuredAliases, name)).sort(),
    ).toEqual(Object.keys(structuredAliases).sort());
    for (const name of Object.keys(structuredAliases)) {
      expect(ts.isTypeAliasDeclaration(declarations.get(name) as ts.Node), name).toBe(true);
    }
  });

  it('generates importable Markdown with exact direct TypeDoc display projections', async () => {
    const temporaryRoot = await mkdtemp(join(tmpdir(), 'tego-sheet-typedoc-'));
    const outputDirectory = join(temporaryRoot, 'api');
    try {
      const typeDocOptions = {
        docsPath: temporaryRoot,
        entryPoints: ['src/index.ts'],
        excludeInternal: true,
        excludePrivate: true,
        excludeProtected: true,
        out: outputDirectory,
        plugin: [
          'typedoc-plugin-markdown',
          'typedoc-docusaurus-theme',
          publicApiProjectionPluginPath,
        ],
        readme: 'none',
        treatValidationWarningsAsErrors: true,
        treatWarningsAsErrors: true,
        tsconfig: 'tsconfig.json',
      } as TypeDocOptions & { docsPath: string };
      const app = await Application.bootstrapWithPlugins(typeDocOptions);
      expect([app.logger.errorCount, app.logger.warningCount], 'TypeDoc bootstrap logger').toEqual([
        0, 0,
      ]);

      const project = await app.convert();
      if (project === undefined) throw new Error('TypeDoc must convert the public entry point');
      expect([app.logger.errorCount, app.logger.warningCount], 'TypeDoc conversion logger').toEqual(
        [0, 0],
      );
      app.validate(project);
      expect([app.logger.errorCount, app.logger.warningCount], 'TypeDoc validation logger').toEqual(
        [0, 0],
      );

      const serialized = app.serializer.projectToObject(project, normalizePath(root));
      const serializedRecord = serialized as unknown as Readonly<Record<string, unknown>>;
      const rootChildren = directRecords(serializedRecord, 'children');
      expect(rootChildren.map((child) => child.name).sort()).toEqual([...expectedPublicExports]);
      expect(rootChildren).toHaveLength(55);

      const aliases = new Map(
        rootChildren
          .filter(
            (child) =>
              typeof child.name === 'string' && Object.hasOwn(structuredAliases, child.name),
          )
          .map((child) => [child.name as string, child]),
      );
      expect([...aliases.keys()].sort()).toEqual(Object.keys(structuredAliases).sort());

      for (const [name, expectedChildren] of Object.entries(structuredAliases)) {
        const reflection = aliases.get(name);
        const children = directRecords(reflection, 'children');
        expect(children.map((child) => child.name).sort(), `${name} direct children`).toEqual([
          ...expectedChildren,
        ]);
        expect(
          children.filter((child) => !hasSummary(child)).map((child) => child.name),
          `${name} direct child summaries`,
        ).toEqual([]);

        const indexSignatures = directRecords(reflection, 'indexSignatures');
        expect(indexSignatures, `${name} direct index signature`).toHaveLength(1);
        expect(summaryText(indexSignatures[0] ?? {}), `${name} index summary`).not.toBe('');
        if (name === 'CellsData') {
          expect(summaryText(indexSignatures[0] ?? {})).toBe(
            'Cell data stored at a zero-based decimal column index.',
          );
        }

        const selfReferences: string[] = [];
        visitRecords(reflection, (record) => {
          if (record.type === 'reference' && record.name === name) selfReferences.push(name);
        });
        expect(selfReferences, `${name} recursive same-name references`).toEqual([]);
      }

      const tegoSheetProps = rootChildren.find((child) => child.name === 'TegoSheetProps');
      const propChildren = directRecords(tegoSheetProps, 'children');
      expect(propChildren.map((child) => child.name).sort()).toEqual([...tegoSheetPropNames]);
      expect(propChildren.filter((child) => !hasSummary(child)).map((child) => child.name)).toEqual(
        [],
      );
      expect(JSON.stringify(tegoSheetProps)).not.toContain('TegoSheetCallbacks');
      expect(
        localReferenceViolations(serializedRecord, localDeclarationNames(createPublicProgram())),
      ).toEqual([]);

      const activeSheetChange = rootChildren.find(
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

      await app.generateOutputs(project);
      expect([app.logger.errorCount, app.logger.warningCount], 'TypeDoc output logger').toEqual([
        0, 0,
      ]);
      const propsPath = await findFile(outputDirectory, 'TegoSheetProps.md');
      if (propsPath === undefined) throw new Error('TypeDoc must generate TegoSheetProps.md');
      const propsMarkdown = await readFile(propsPath, 'utf8');
      expect(propsMarkdown).not.toContain('TegoSheetCallbacks');
      expect(propsMarkdown).not.toMatch(/Inherited from/iu);
      for (const callback of callbackNames) {
        expect(propsMarkdown, `${callback} Markdown heading`).toMatch(
          new RegExp(`^#+ ${callback}\\??(?:\\(.*\\))?$`, 'mu'),
        );
      }

      const cellsPath = await findFile(outputDirectory, 'CellsData.md');
      if (cellsPath === undefined) throw new Error('TypeDoc must generate CellsData.md');
      const cellsMarkdown = await readFile(cellsPath, 'utf8');
      expect(cellsMarkdown).toContain('Cell data stored at a zero-based decimal column index.');
    } finally {
      await rm(temporaryRoot, { force: true, recursive: true });
    }
  }, 30_000);
});
