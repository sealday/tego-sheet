import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const consumer = process.env.TEGO_SHEET_CONSUMER;
assert.ok(consumer, 'TEGO_SHEET_CONSUMER must point at the clean installed fixture');
const expectedExports = [
  '.',
  './styles.css',
  './locales/en',
  './locales/de',
  './locales/nl',
  './locales/zh-cn',
  './interchange',
  './analysis',
  './sdk',
  './integrations',
  './output/pdf',
  './output/xlsx',
  './output/image',
  './package.json',
];

test('the analysis subpath exposes Worker-safe analysis protocols', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const analysis = await import('tego-sheet/analysis');
      const expected = [
        'buildSlicerValueIndex',
        'chartAffectedByChanges',
        'chartToDisplayCommands',
        'compileSlicerFilterContext',
        'compileSolverModel',
        'createFormulaGoalSeekApplyProposal',
        'createStructuredTableResolver',
        'executeStructuredTableView',
        'planStructuredTableAutoExpand',
        'refreshPivot',
        'resolveChart',
        'resolveSparkline',
        'runSolver',
        'solveFormulaGoalSeek',
        'sparklineAffectedByChanges',
        'sparklineToDisplayCommands'
      ];
      if (JSON.stringify(Object.keys(analysis)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected analysis exports: ' + Object.keys(analysis).join(','));
      }
      if ('document' in globalThis) throw new Error('analysis created a DOM global');
      `,
    ],
    { cwd: consumer },
  );
});

test('the package exposes only the approved root and locale entry points', () => {
  const packageJson = JSON.parse(
    readFileSync(join(consumer, 'node_modules/tego-sheet/package.json'), 'utf8'),
  );
  assert.deepEqual(Object.keys(packageJson.exports), expectedExports);
  assert.equal(packageJson.exports['./locales'], undefined);
});

test('the built root has only the approved runtime exports and internal subpaths are blocked', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const root = await import('tego-sheet');
      const expected = [
        'AiProposalPanel',
        'BUILTIN_FORMULA_COMPATIBILITY',
        'BUILTIN_NUMBER_FORMAT_COMPATIBILITY',
        'BrowserPrintError',
          'ConditionalFormatError',
          'DataTransformError',
          'FormulaNameConflictError',
          'FormulaSyntaxError',
        'InterchangeError',
        'IsolatedBrowserPrintAdapter',
        'NumberFormatSyntaxError',
        'SheetObjectError',
        'TEMPLATE_COMPILER_VERSION',
        'TegoSheet',
        'TegoSheetException',
        'TemplateDesigner',
        'TemplateExpressionError',
        'TemplatePreview',
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
        'createStructuredTableResolver',
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
        'planStructuredTableAutoExpand',
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
      ];
      if (JSON.stringify(Object.keys(root)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected root exports: ' + Object.keys(root).join(','));
      }
      const exception = new root.TegoSheetException({
        code: 'INVALID_COMMAND', message: 'probe', recoverable: false,
      });
      if (!(exception instanceof root.TegoSheetException)) throw new Error('bad exception runtime');
      const document = root.createSpreadsheetDocument({
        id: 'packed-document', sheetId: 'packed-sheet',
      });
      const serialized = root.serializeSpreadsheetDocument(document);
      const parsed = root.parseSpreadsheetDocument(serialized);
      if (!parsed.ok || parsed.document.id !== 'packed-document') {
        throw new Error('bad Workbook 2.0 runtime');
      }
      const migrated = root.migrateLegacyWorkbook(
        { name: 'Legacy' },
        {
          ids: {
            documentId: () => 'migrated-document',
            sheetId: (index) => 'migrated-sheet-' + index,
          },
        },
      );
      if (!migrated.ok || migrated.document.id !== 'migrated-document') {
        throw new Error('bad legacy migration runtime');
      }
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );

  for (const subpath of [
    'tego-sheet/locales',
    'tego-sheet/core',
    'tego-sheet/engine',
    'tego-sheet/react/tego-sheet',
    'tego-sheet/src/index',
    'tego-sheet/legacy',
  ]) {
    const result = spawnSync(
      process.execPath,
      ['--input-type=module', '--eval', `await import(${JSON.stringify(subpath)})`],
      { cwd: consumer, encoding: 'utf8' },
    );
    assert.notEqual(result.status, 0, `${subpath} unexpectedly resolved`);
    assert.match(result.stderr, /ERR_PACKAGE_PATH_NOT_EXPORTED/);
  }
});

test('the optional PDF subpath exposes only its adapter contract', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const pdf = await import('tego-sheet/output/pdf');
      const expected = ['OutputAdapterError', 'PdfAdapter'];
      if (JSON.stringify(Object.keys(pdf)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected PDF exports: ' + Object.keys(pdf).join(','));
      }
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('the optional XLSX subpath exposes only its adapter contract', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const xlsx = await import('tego-sheet/output/xlsx');
      const expected = ['OutputAdapterError', 'XlsxAdapter'];
      if (JSON.stringify(Object.keys(xlsx)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected XLSX exports: ' + Object.keys(xlsx).join(','));
      }
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('the optional image subpath exposes only its adapter contract', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const image = await import('tego-sheet/output/image');
      const expected = ['ImageAdapter', 'OutputAdapterError'];
      if (JSON.stringify(Object.keys(image)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected image exports: ' + Object.keys(image).join(','));
      }
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('the interchange subpath exposes only Worker-safe reader and writer contracts', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const interchange = await import('tego-sheet/interchange');
      const expected = [
        'InterchangeError',
        'createCsvReader',
        'createCsvWriter',
          'createOdsReader',
          'createOdsWriter',
        'createTsvReader',
        'createTsvWriter',
          'createXlsxReader',
          'createXlsxWriter',
      ];
      if (JSON.stringify(Object.keys(interchange)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected interchange exports: ' + Object.keys(interchange).join(','));
      }
      if ('document' in globalThis) throw new Error('interchange created a DOM global');
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('the SDK subpath exposes only approved public extension runtimes', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const sdk = await import('tego-sheet/sdk');
      const expected = [
        'ADAPTER_KINDS',
        'AdapterSdkError',
        'DEFAULT_ADAPTER_SCOPE_LIMITS',
        'TemplateModuleSdkError',
        'createAdapterRegistry',
        'createCapabilityGrant',
        'createCellEditorSession',
        'createCellTypeRegistry',
        'createTemplateModuleRegistry',
        'executeTemplateModulePipeline',
        'resolveCustomCell',
      ];
      if (JSON.stringify(Object.keys(sdk)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected SDK exports: ' + Object.keys(sdk).join(','));
      }
      if ('document' in globalThis) throw new Error('SDK created a DOM global');
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('the integrations subpath exposes only host-owned protocol runtimes', () => {
  execFileSync(
    process.execPath,
    [
      '--input-type=module',
      '--eval',
      `
      const integrations = await import('tego-sheet/integrations');
      const expected = [
        'checkpointDocumentVersion',
        'createAiProposalSession',
        'createCollaborationOutboxCoordinator',
        'createCollaborationSession',
        'createCommentAnchorOutboxCoordinator',
        'createCommentStore',
        'createControllerAiProposalSession',
        'createControllerRemoteTransactionBoundary',
        'createPermissionSnapshot',
        'createPermissionStore',
        'createPersistenceController',
        'createPersistenceSession',
        'createPresenceStore',
        'createRemoteOperationProcessor',
        'createRestoreVersionProposal',
        'createWorkbookTransactionPermissionGate',
        'deriveWorkbookCommandPermissionRequests',
        'diffDocumentVersions',
        'diffDocumentVersionsAsync',
        'evaluatePermission',
        'listDocumentVersions',
        'loadHistoryPreview',
        'loadPersistedDocument',
        'projectAiContext',
        'projectCommentPrintContent',
        'restoreDocumentVersion',
        'sanitizeCommentRichText',
        'summarizeAiContext',
        'transformCommentAnchor',
      ];
      if (JSON.stringify(Object.keys(integrations)) !== JSON.stringify(expected)) {
        throw new Error('Unexpected integrations exports: ' + Object.keys(integrations).join(','));
      }
    `,
    ],
    { cwd: consumer, stdio: 'pipe' },
  );
});

test('each locale entry exports only its intended dictionary', () => {
  const probes = [
    ['en', 'en', 'en'],
    ['de', 'de', 'de'],
    ['nl', 'nl', 'nl'],
    ['zh-cn', 'zhCN', 'zh-CN'],
  ];
  for (const [subpath, name, id] of probes) {
    execFileSync(
      process.execPath,
      [
        '--input-type=module',
        '--eval',
        `
      const locale = await import(${JSON.stringify(`tego-sheet/locales/${subpath}`)});
      if (JSON.stringify(Object.keys(locale)) !== JSON.stringify([${JSON.stringify(name)}])) {
        throw new Error('Unexpected locale exports: ' + Object.keys(locale).join(','));
      }
      if (locale[${JSON.stringify(name)}].id !== ${JSON.stringify(id)}) throw new Error('bad id');
    `,
      ],
      { cwd: consumer, stdio: 'pipe' },
    );
  }
});
