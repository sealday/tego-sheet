import { expect, it } from 'vitest';

it('imports the library entry without browser globals', async () => {
  const entry = await import('../../src/index');

  expect(Object.getOwnPropertyNames(entry)).toEqual([
    'TegoSheetException',
    'createSpreadsheetDocument',
    'migrateLegacyWorkbook',
    'parseSpreadsheetDocument',
    'serializeSpreadsheetDocument',
    'createDocumentController',
    'BUILTIN_FORMULA_COMPATIBILITY',
    'FormulaSyntaxError',
    'createFormulaEngine',
    'createFormulaFunctionRegistry',
    'parseFormula',
    'renderFormula',
    'resolveFormulaReferences',
    'translateFormula',
    'BUILTIN_NUMBER_FORMAT_COMPATIBILITY',
    'NumberFormatSyntaxError',
    'createNumberFormatter',
    'parseNumberFormat',
    'renderNumberFormatToken',
    'createFontMetrics',
    'createPresentationCache',
    'createPresentationResolver',
    'createPrintDisplayList',
    'TegoSheet',
  ]);
});
