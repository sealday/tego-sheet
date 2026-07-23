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
  './output/pdf',
  './output/xlsx',
  './output/image',
  './package.json',
];

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
        'BUILTIN_FORMULA_COMPATIBILITY',
        'BUILTIN_NUMBER_FORMAT_COMPATIBILITY',
        'BrowserPrintError',
        'ConditionalFormatError',
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
        'evaluateTemplateExpression',
        'executeValidatedCellEdit',
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
        'createTsvReader',
        'createTsvWriter',
        'createXlsxReader',
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
