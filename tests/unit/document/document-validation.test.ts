import { describe, expect, it } from 'vitest';
import { parseSpreadsheetDocument } from '../../../src/document';
import type {
  DocumentDiagnosticCode,
  DocumentParseOptions,
  SpreadsheetDocumentInput,
} from '../../../src/document';

const validDocument = (): SpreadsheetDocumentInput => ({
  schemaVersion: 2,
  id: 'document-validation',
  workbook: {
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        cells: [],
        merges: [],
      },
    ],
    styles: [{ id: 'style-1', value: { fontWeight: 'bold' } }],
    validations: [{ id: 'validation-1', value: { kind: 'list' } }],
    settings: { dateSystem: 'excel-1900' },
  },
  templates: [
    {
      id: 'template-1',
      name: 'Print template',
      sheetId: 'sheet-1',
      range: {
        sheetId: 'sheet-1',
        start: { row: 0, column: 0 },
        end: { row: 2, column: 2 },
      },
      printProfile: {
        paperSize: 'A4',
        orientation: 'portrait',
        margins: { top: 1, right: 1, bottom: 1, left: 1 },
      },
    },
  ],
  resources: {
    items: [{ id: 'resource-1', kind: 'image', mimeType: 'image/png', byteLength: 12 }],
  },
  extensions: { 'vendor.example': { enabled: true } },
});

const codesOf = (
  input: unknown,
  options?: DocumentParseOptions,
): readonly DocumentDiagnosticCode[] => {
  const result = parseSpreadsheetDocument(input, options);
  expect(result.ok).toBe(false);
  if (result.ok) throw new Error('expected invalid document');
  expect('document' in result).toBe(false);
  return result.diagnostics.map(({ code }) => code);
};

describe('Workbook 2.0 validation', () => {
  it('rejects unsupported schema versions atomically', () => {
    expect(codesOf({ ...validDocument(), schemaVersion: 3 })).toContain(
      'UNSUPPORTED_SCHEMA_VERSION',
    );
  });

  it('reports malformed document shapes with a stable code and structured path', () => {
    const result = parseSpreadsheetDocument({ schemaVersion: 2, workbook: null });
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('expected invalid document');

    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'DOCUMENT_SCHEMA_INVALID',
          domain: 'document',
          stage: 'decode',
          details: expect.objectContaining({ path: '$.workbook' }),
        }),
      ]),
    );
  });

  it('aggregates duplicate IDs across stable-ID collections', () => {
    const fixture = validDocument();
    fixture.workbook.sheets.push({ ...fixture.workbook.sheets[0]!, name: 'Duplicate' });
    fixture.workbook.styles.push({ id: 'style-1', value: null });

    expect(codesOf(fixture).filter((code) => code === 'DUPLICATE_ID')).toHaveLength(2);
  });

  it('aggregates dangling style, validation, resource, template, and sheet references', () => {
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.cells.push({
      row: 0,
      column: 0,
      cell: {
        input: { type: 'string', value: 'references' },
        styleId: 'missing-style',
        validationId: 'missing-validation',
        resourceId: 'missing-resource',
        templateId: 'missing-template',
      },
    });
    fixture.templates[0]!.sheetId = 'missing-sheet';

    expect(codesOf(fixture).filter((code) => code === 'DANGLING_REFERENCE')).toHaveLength(5);
  });

  it('reports invalid and cross-sheet ranges', () => {
    const fixture = validDocument();
    fixture.templates[0]!.range = {
      sheetId: 'sheet-2',
      start: { row: 3, column: 0 },
      end: { row: 1, column: 2 },
    };

    expect(codesOf(fixture)).toContain('INVALID_RANGE');
  });

  it('reports overlapping merges', () => {
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.merges.push(
      {
        start: { row: 0, column: 0 },
        end: { row: 2, column: 2 },
      },
      {
        start: { row: 1, column: 1 },
        end: { row: 3, column: 3 },
      },
    );

    expect(codesOf(fixture)).toContain('INVALID_MERGE');
  });

  it('rejects duplicate sparse coordinates atomically', () => {
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.cells.push(
      { row: 2, column: 1, cell: { input: { type: 'blank' } } },
      { row: 2, column: 1, cell: { input: { type: 'string', value: 'duplicate' } } },
    );

    const result = parseSpreadsheetDocument(fixture);
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('expected invalid document');
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'DOCUMENT_SCHEMA_INVALID',
        details: { path: '$.workbook.sheets[0].cells[1]' },
      }),
    );
  });

  it('reports diagnostic paths against caller order before canonical sorting', () => {
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.cells.push(
      { row: 9, column: 0, cell: { input: { type: 'blank' } } },
      {
        row: 1,
        column: 0,
        cell: { input: { type: 'blank' }, styleId: 'missing-style' },
      },
    );

    const result = parseSpreadsheetDocument(fixture);
    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('expected invalid document');
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'DANGLING_REFERENCE',
        details: { path: '$.workbook.sheets[0].cells[1].cell.styleId' },
      }),
    );
  });

  it.each([
    { 'vendor.example': undefined },
    { 'vendor.example': Number.NaN },
    { 'vendor.example': () => undefined },
    { unnamespaced: true },
  ])('rejects invalid or non-namespaced extension data: %o', (extensions) => {
    expect(codesOf({ ...validDocument(), extensions })).toContain('INVALID_EXTENSION_DATA');
  });

  it.each([
    [{ maxSheets: 0 }, 'sheet'],
    [{ maxBytes: 1 }, 'byte'],
  ] as const)('enforces configured %s limits', (limits, _label) => {
    expect(codesOf(validDocument(), { limits })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('enforces the configured non-empty cell limit', () => {
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.cells.push({
      row: 0,
      column: 0,
      cell: { input: { type: 'blank' } },
    });

    expect(codesOf(fixture, { limits: { maxCells: 0 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('aggregates independent diagnostics without exposing a partial document', () => {
    const fixture = validDocument();
    fixture.workbook.sheets.push({ ...fixture.workbook.sheets[0]! });
    fixture.workbook.sheets[0]!.cells.push({
      row: -1,
      column: 0,
      cell: {
        input: { type: 'string', value: 'bad' },
        styleId: 'missing-style',
      },
    });
    (fixture as { extensions: Record<string, unknown> }).extensions = { invalid: undefined };

    const codes = codesOf(fixture);
    expect(new Set(codes)).toEqual(
      expect.objectContaining(
        new Set(['DUPLICATE_ID', 'DOCUMENT_SCHEMA_INVALID', 'INVALID_EXTENSION_DATA']),
      ),
    );
  });

  it('enforces string byte limits before attempting JSON decoding', () => {
    const result = parseSpreadsheetDocument('not valid JSON', { limits: { maxBytes: 1 } });
    expect(result).toMatchObject({
      ok: false,
      diagnostics: [{ code: 'DOCUMENT_LIMIT_EXCEEDED', details: { path: '$' } }],
    });
  });
});
