import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ts from 'typescript';
import { Application, normalizePath } from 'typedoc';
import type { TypeDocOptions } from 'typedoc';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const entryPoint = join(root, 'src/index.ts');

const expectedPublicExports = [
  'ActiveSheetChangeEvent',
  'AddSheetCommand',
  'AdvancedCompileOptions',
  'AdvancedExpansionResult',
  'AdvancedFormulaBindingContext',
  'AdvancedRepeatBase',
  'AdvancedValidationRequest',
  'AdvancedValidationResult',
  'AdvancedValidationRule',
  'ApplyFilterViewInput',
  'AutoFilterData',
  'AutoFilterItemData',
  'AutoFilterSortData',
  'AutofillCommand',
  'AutofillTransform',
  'BUILTIN_FORMULA_COMPATIBILITY',
  'BUILTIN_NUMBER_FORMAT_COMPATIBILITY',
  'BindingId',
  'BorderLine',
  'BorderMode',
  'BrowserPrintCleanupReason',
  'BrowserPrintError',
  'BrowserPrintErrorCode',
  'BrowserPrintOptions',
  'BrowserPrintResult',
  'BrowserPrintSvgPage',
  'CalculationEnvironment',
  'CalculationResult',
  'Cell',
  'CellAddress',
  'CellBorders',
  'CellData',
  'CellEditEvent',
  'CellInput',
  'CellPoint',
  'CellPresentation',
  'CellRange',
  'CellStyle',
  'CellsData',
  'ChangeSource',
  'ClearContentsCommand',
  'ClearFilterCommand',
  'ClearFormatCommand',
  'ColsData',
  'ColumnData',
  'CompilationResult',
  'CompiledTemplate',
  'CompiledTemplateExpression',
  'ConditionalCellRule',
  'ConditionalColorScale',
  'ConditionalEffect',
  'ConditionalEvaluationInput',
  'ConditionalEvaluationResult',
  'ConditionalExpression',
  'ConditionalFormat',
  'ConditionalFormatError',
  'ConditionalFormatLimits',
  'ConditionalRangeBinding',
  'ConditionalRule',
  'ConditionalRuleId',
  'ConditionalStyle',
  'ConditionalStylePatch',
  'CreateDocumentOptions',
  'DataAnomalyAnalysis',
  'DataAnomalyAnalysisOptions',
  'DataAnomalyAnalysisRequest',
  'DataAnomalyCheck',
  'DataToolPreviewContext',
  'DataTransform',
  'DataTransformCommitResult',
  'DataTransformError',
  'DataTransformPreview',
  'DataTransformPreviewOptions',
  'DecodedResourceImage',
  'DeleteSheetCommand',
  'DelimitedWriteOptions',
  'DependencyChange',
  'Diagnostic',
  'DiagnosticDomain',
  'DiagnosticLocation',
  'DiagnosticStage',
  'DisplayRect',
  'DocumentCellAddress',
  'DocumentCellPoint',
  'DocumentCellRange',
  'DocumentChange',
  'DocumentChangedRange',
  'DocumentCommand',
  'DocumentCommandEnvelope',
  'DocumentCommittedTransaction',
  'DocumentController',
  'DocumentControllerEvent',
  'DocumentControllerOptions',
  'DocumentControllerSnapshot',
  'DocumentDiagnostic',
  'DocumentDiagnosticCode',
  'DocumentExecuteOptions',
  'DocumentId',
  'DocumentLimits',
  'DocumentParseOptions',
  'DocumentParseResult',
  'DocumentPatchOperation',
  'DocumentSheetChange',
  'DocumentSheetId',
  'DocumentTransactionDiagnostic',
  'DocumentTransactionEnvelope',
  'DocumentTransactionOptions',
  'DocumentTransactionPermissionContext',
  'DocumentTransactionPermissionGate',
  'DocumentTransactionPreview',
  'DocumentTransactionResult',
  'ExtensionStore',
  'FillSeriesTransform',
  'FilterDefinition',
  'FilterView',
  'FilterViewPredicate',
  'FilterViewSession',
  'FindReplaceTransform',
  'FontFaceMetrics',
  'FontMetrics',
  'FontMetricsOptions',
  'FontStyle',
  'FormatContext',
  'FormulaAddress',
  'FormulaAst',
  'FormulaBoundReference',
  'FormulaDependencyGraph',
  'FormulaDiagnostic',
  'FormulaEngine',
  'FormulaEngineOptions',
  'FormulaError',
  'FormulaFunctionCompatibility',
  'FormulaFunctionContext',
  'FormulaFunctionDefinition',
  'FormulaFunctionRegistry',
  'FormulaNameConflictError',
  'FormulaNameDefinition',
  'FormulaNameRegistry',
  'FormulaNodeBase',
  'FormulaProgram',
  'FormulaReference',
  'FormulaSpillError',
  'FormulaSpillPlan',
  'FormulaSyntaxError',
  'FormulaTableBindingRequest',
  'FormulaTableBindingResolver',
  'FormulaTableBindingResult',
  'FormulaTranslation',
  'FormulaValue',
  'GeneratedCalculatedCell',
  'GeneratedConditionalCellRule',
  'GeneratedConditionalColorScale',
  'GeneratedConditionalFormat',
  'GeneratedConditionalStyle',
  'GeneratedDocument',
  'GeneratedDocumentForBrowserPrint',
  'GeneratedPrintPage',
  'GeneratedWorksheet',
  'GroupCommand',
  'GroupId',
  'HideColumnCommand',
  'HideRowCommand',
  'HistoryCommand',
  'HorizontalAlign',
  'IndexedSheetCommand',
  'InterchangeError',
  'InterchangeErrorCode',
  'InterchangeFormat',
  'InterchangeInput',
  'InterchangeLimits',
  'InterchangeReadOptions',
  'InterchangeSecurityReport',
  'InterchangeWriteOptions',
  'IsolatedBrowserPrintAdapter',
  'IsolatedBrowserPrintAdapterOptions',
  'JsonValue',
  'LegacyMigrationDiagnosticCode',
  'LegacyMigrationIdFactory',
  'LegacyMigrationOptions',
  'LocaleDefinition',
  'LocaleMessages',
  'MergeCommand',
  'MigrationDiagnostic',
  'MigrationResult',
  'NumberFormatAst',
  'NumberFormatCompatibility',
  'NumberFormatCondition',
  'NumberFormatConditionOperator',
  'NumberFormatSection',
  'NumberFormatSyntaxError',
  'NumberFormatToken',
  'NumberFormatter',
  'ObjectAnchor',
  'ObjectBase',
  'ObjectCoordinateTransform',
  'ObjectDisplayContext',
  'ObjectGeometry',
  'ObjectId',
  'ObjectOffset',
  'ObjectRect',
  'ObjectRepeatPolicy',
  'PageBand',
  'PageBreak',
  'PageSetup',
  'PaintFormatCommand',
  'PaperDefinition',
  'PasteEvent',
  'PasteExternalCommand',
  'PasteInternalCommand',
  'PasteMode',
  'PresentationAnnotation',
  'PresentationCache',
  'PresentationCacheOptions',
  'PresentationCacheStats',
  'PresentationEnvironment',
  'PresentationProblem',
  'PresentationResolver',
  'PresentationResolverOptions',
  'PresentationRevisions',
  'PresentationValid',
  'PresentationValidation',
  'PrintClipCommand',
  'PrintDisplayCell',
  'PrintDisplayCommand',
  'PrintDisplayList',
  'PrintDisplayListInput',
  'PrintDisplayPage',
  'PrintDisplayPageInput',
  'PrintDocument',
  'PrintFillRectCommand',
  'PrintGroupCommand',
  'PrintImageCommand',
  'PrintLineCommand',
  'PrintLinkCommand',
  'PrintMargins',
  'PrintPathCommand',
  'PrintProfile',
  'PrintScale',
  'PrintStrokeRectCommand',
  'PrintTarget',
  'PrintTextCommand',
  'QrResourceOptions',
  'RedoCommand',
  'ReferenceResolutionResult',
  'RegistryEntry',
  'RemoveConditionalFormatCommand',
  'RemoveDuplicatesTransform',
  'RemoveFilterViewCommand',
  'RemoveSheetObjectCommand',
  'RemoveValidationCommand',
  'RenameSheetCommand',
  'RenderEnvironment',
  'RenderLimits',
  'RenderRequest',
  'RenderResult',
  'RepeatColumnsBinding',
  'RepeatPageBinding',
  'RepeatRangeBinding',
  'RepeatRowsBinding',
  'RepeatSheetBinding',
  'RepeatedObjectRef',
  'ResizeColumnCommand',
  'ResizeRowCommand',
  'ResolveContext',
  'ResolvedFontMetrics',
  'ResolvedResource',
  'ResolvedResourceCache',
  'ResolvedResourceStore',
  'ResolvedResourceVector',
  'ResolvedStyle',
  'ResourceCapabilityRegistry',
  'ResourceFontHandle',
  'ResourceId',
  'ResourceKernelEnvironment',
  'ResourceLimits',
  'ResourceMetadata',
  'ResourcePipelineOptions',
  'ResourcePurpose',
  'ResourceRef',
  'ResourceResolutionResult',
  'ResourceResolver',
  'ResourceResolverRegistry',
  'ResourceStore',
  'ResourceType',
  'RowData',
  'RowsData',
  'ScalarFormulaValue',
  'Selection',
  'SetBorderCommand',
  'SetCellInputCommand',
  'SetCellMetadataCommand',
  'SetCellTextCommand',
  'SetConditionalFormatCommand',
  'SetFilterCommand',
  'SetFilterViewCommand',
  'SetFreezeCommand',
  'SetSheetObjectCommand',
  'SetStyleCommand',
  'SetValidationCommand',
  'Sheet',
  'SheetColumn',
  'SheetColumnOptions',
  'SheetData',
  'SheetFilter',
  'SheetFilterItem',
  'SheetGroup',
  'SheetId',
  'SheetObject',
  'SheetObjectError',
  'SheetOptions',
  'SheetRange',
  'SheetRow',
  'SheetRowOptions',
  'SheetTabItem',
  'SheetTabsRenderProps',
  'SheetTabsRenderer',
  'SortCommand',
  'SourceSpan',
  'SparseCell',
  'SplitTextTransform',
  'SpreadsheetDocument',
  'SpreadsheetTemplate',
  'StoredSpreadsheetTemplate',
  'StructuralMapping',
  'StructuralObjectMapping',
  'StyleId',
  'StyleRegistry',
  'SubtemplateBinding',
  'TEMPLATE_COMPILER_VERSION',
  'TegoSheet',
  'TegoSheetError',
  'TegoSheetErrorCode',
  'TegoSheetException',
  'TegoSheetHandle',
  'TegoSheetProps',
  'TemplateBinding',
  'TemplateDesigner',
  'TemplateDesignerProps',
  'TemplateExpressionError',
  'TemplateExpressionNode',
  'TemplateExpressionScope',
  'TemplateFormatter',
  'TemplateFormatterRegistry',
  'TemplateIR',
  'TemplateIRBinding',
  'TemplateId',
  'TemplatePreview',
  'TemplatePreviewProps',
  'TemplatePrintProfile',
  'TemplateRegionNode',
  'TemplateResourceBinding',
  'ToggleGroupCommand',
  'ToolbarAction',
  'ToolbarRenderProps',
  'ToolbarRenderer',
  'TransactionChangeAggregate',
  'TransactionSheetChange',
  'UndoCommand',
  'UngroupCommand',
  'UnverifiedResource',
  'ValidatedCellEditRequest',
  'ValidatedCellEditResult',
  'ValidatedTransactionRequest',
  'ValidationComparison',
  'ValidationComparisonOperator',
  'ValidationData',
  'ValidationEngine',
  'ValidationEngineOptions',
  'ValidationFormulaContext',
  'ValidationFormulaEvaluator',
  'ValidationId',
  'ValidationIssue',
  'ValidationListSource',
  'ValidationOperator',
  'ValidationRangeComparison',
  'ValidationRegistry',
  'ValidationResolver',
  'ValidationResolverContext',
  'ValidationResolverRegistry',
  'ValidationResult',
  'ValidationRule',
  'ValidationRuleBase',
  'ValidationScalarComparison',
  'ValidationType',
  'ValueBinding',
  'VerticalAlign',
  'Workbook',
  'WorkbookChange',
  'WorkbookChangeKind',
  'WorkbookCommand',
  'WorkbookExportResult',
  'WorkbookImportResult',
  'WorkbookReader',
  'WorkbookSettings',
  'WorkbookWriter',
  'WorksheetVisibility',
  'analyzeDataAnomalies',
  'applyFilterView',
  'bindAdvancedFormula',
  'compileSpreadsheetTemplate',
  'compileTemplateExpression',
  'createBlobResourceResolver',
  'createConditionalFormatEvaluator',
  'createCsvReader',
  'createCsvWriter',
  'createDataTransformPlanner',
  'createDataUrlResourceResolver',
  'createDocumentController',
  'createFilterViewSession',
  'createFontMetrics',
  'createFormulaEngine',
  'createFormulaFunctionRegistry',
  'createFormulaNameRegistry',
  'createNumberFormatter',
  'createOdsReader',
  'createOdsWriter',
  'createPresentationCache',
  'createPresentationResolver',
  'createPrintDisplayList',
  'createResolvedResourceCache',
  'createResourceResolverRegistry',
  'createResourceResolverRegistryFromKernel',
  'createSpreadsheetDocument',
  'createTsvReader',
  'createTsvWriter',
  'createValidationEngine',
  'createValidationResolverRegistry',
  'createXlsxReader',
  'createXlsxWriter',
  'evaluateTemplateExpression',
  'executeValidatedCellEdit',
  'executeValidatedTransaction',
  'expandAdvancedTemplate',
  'hashSpreadsheetDocument',
  'migrateLegacyWorkbook',
  'objectToDisplayCommands',
  'parseFormula',
  'parseNumberFormat',
  'parseSpreadsheetDocument',
  'planFormulaSpill',
  'renderFormula',
  'renderNumberFormatToken',
  'renderSpreadsheetTemplate',
  'resolveFormulaReferences',
  'resolveObjectAnchor',
  'resolveTemplateResources',
  'serializeGeneratedDocumentSvgPages',
  'serializeSpreadsheetDocument',
  'transformObjectAnchor',
  'translateFormula',
  'validatePrintDisplayCommands',
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
  'activePrintProfileId',
  'className',
  'confirmValidationWarning',
  'defaultDocument',
  'document',
  'initialActiveSheetIndex',
  'locale',
  'mode',
  'onActivePrintProfileChange',
  'onActiveSheetChange',
  'onCellEdit',
  'onDiagnostics',
  'onDocumentChange',
  'onError',
  'onPaste',
  'onSelectionChange',
  'onTemplateChange',
  'options',
  'readOnly',
  'renderEnvironment',
  'sampleData',
  'sheetTabs',
  'style',
  'template',
  'toolbar',
  'validationEngine',
] as const;

const callbackNames = [
  'onActivePrintProfileChange',
  'onActiveSheetChange',
  'onCellEdit',
  'onDiagnostics',
  'onDocumentChange',
  'onError',
  'onPaste',
  'onSelectionChange',
  'onTemplateChange',
] as const;

const sparseCollectionNames = ['CellsData', 'ColsData', 'RowsData'] as const;
const sparseCollectionSummary = 'JSON-compatible entry stored at a sparse decimal index.';

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

const publicProgram = createPublicProgram();
const publicChecker = publicProgram.getTypeChecker();
const publicLocalDeclarationNames = localDeclarationNames(publicProgram);

describe('public API documentation', () => {
  it('documents the exact root export surface and its public members', () => {
    const source = publicProgram.getSourceFile(entryPoint);
    if (source === undefined)
      throw new Error('src/index.ts must be part of the TypeScript program');
    const moduleSymbol = publicChecker.getSymbolAtLocation(source);
    if (moduleSymbol === undefined) throw new Error('src/index.ts must be a module');
    const exports = publicChecker
      .getExportsOfModule(moduleSymbol)
      .sort((left, right) => (left.name < right.name ? -1 : left.name > right.name ? 1 : 0));

    expect(exports.map((symbol) => symbol.name)).toEqual(expectedPublicExports);

    const undocumented = exports
      .flatMap((symbol) => publicDeclarations(symbol, publicChecker))
      .filter((declaration) => !hasDocumentation(declaration, publicChecker))
      .map((declaration) => declaration.name);

    expect(undocumented).toEqual([]);
  });

  it('keeps all 14 structured compiler declarations as type aliases', () => {
    const source = publicProgram.getSourceFile(join(root, 'src/core/types/workbook.ts'));
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
        plugin: ['typedoc-plugin-markdown', 'typedoc-docusaurus-theme'],
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
      expect(rootChildren).toHaveLength(expectedPublicExports.length);

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
        if (sparseCollectionNames.includes(name as (typeof sparseCollectionNames)[number])) {
          expect(summaryText(indexSignatures[0] ?? {}), `${name} exact index summary`).toBe(
            sparseCollectionSummary,
          );
        }

        const selfReferences: string[] = [];
        visitRecords(reflection, (record) => {
          if (record.type === 'reference' && record.name === name) selfReferences.push(name);
        });
        expect(selfReferences, `${name} recursive same-name references`).toEqual([]);
      }

      const tegoSheetProps = rootChildren.find((child) => child.name === 'TegoSheetProps');
      const propChildren: Readonly<Record<string, unknown>>[] = [];
      visitRecords(tegoSheetProps?.type, (record) => {
        if (
          record.kind === 1024 &&
          typeof record.name === 'string' &&
          tegoSheetPropNames.includes(record.name as (typeof tegoSheetPropNames)[number])
        ) {
          propChildren.push(record);
        }
      });
      expect([...new Set(propChildren.map((child) => child.name))].sort()).toEqual([
        ...tegoSheetPropNames,
      ]);
      for (const name of tegoSheetPropNames) {
        expect(
          propChildren.some((child) => child.name === name && hasSummary(child)),
          `${name} has at least one documented ownership branch`,
        ).toBe(true);
      }
      expect(JSON.stringify(tegoSheetProps)).not.toContain('TegoSheetCallbacks');
      expect(localReferenceViolations(serializedRecord, publicLocalDeclarationNames)).toEqual([]);

      const activeSheetChange = rootChildren.find(
        (child) => child.name === 'ActiveSheetChangeEvent',
      );
      const activeSheet = directRecords(activeSheetChange, 'children').find(
        (value) => value.name === 'sheet',
      );
      expect(activeSheet).toBeDefined();
      expect(summaryText(activeSheet ?? {})).toBe(
        'Worksheet reported as active by the activation event.',
      );

      await app.generateOutputs(project);
      expect([app.logger.errorCount, app.logger.warningCount], 'TypeDoc output logger').toEqual([
        0, 0,
      ]);
      const propsMarkdown = await readFile(
        join(outputDirectory, 'type-aliases/TegoSheetProps.md'),
        'utf8',
      );
      expect(propsMarkdown).not.toContain('TegoSheetCallbacks');
      expect(propsMarkdown).not.toMatch(/Inherited from/iu);
      for (const callback of callbackNames) {
        expect(propsMarkdown, `${callback} Markdown heading`).toMatch(
          new RegExp(`^#+ ${callback}\\??(?:\\(.*\\))?$`, 'mu'),
        );
      }

      const sparseCollectionMarkdown = await Promise.all(
        sparseCollectionNames.map(async (name) => [
          name,
          await readFile(join(outputDirectory, `interfaces/${name}.md`), 'utf8'),
        ]),
      );
      for (const [name, markdown] of sparseCollectionMarkdown) {
        expect(markdown, `${name} Markdown index summary`).toContain(sparseCollectionSummary);
      }
    } finally {
      await rm(temporaryRoot, { force: true, recursive: true });
    }
  }, 30_000);
});
