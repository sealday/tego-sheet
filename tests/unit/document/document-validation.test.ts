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
      bindings: [
        {
          id: 'binding-1' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
          expression: 'value',
        },
      ],
      printProfiles: [
        {
          id: 'profile-1',
          name: 'A4',
          targets: [{ type: 'sheet', sheetId: 'sheet-1' as never }],
          page: {
            paper: { type: 'A4' },
            orientation: 'portrait',
            margins: { top: 1, right: 1, bottom: 1, left: 1 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: true,
          showHeadings: false,
        },
      ],
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

function documentWithWorkbookGetter(
  decodedWorkbook: SpreadsheetDocumentInput['workbook'],
): SpreadsheetDocumentInput {
  const fixture = validDocument();
  const measuredWorkbook = fixture.workbook;
  let reads = 0;
  Object.defineProperty(fixture, 'workbook', {
    enumerable: true,
    get: () => {
      reads += 1;
      return reads < 5 ? measuredWorkbook : decodedWorkbook;
    },
  });
  return fixture;
}

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
    fixture.templates[0]!.bindings[0] = {
      id: 'binding-1' as never,
      type: 'value',
      target: { sheetId: 'missing-sheet' as never, row: 0, column: 0 },
      expression: 'value',
    };

    expect(codesOf(fixture).filter((code) => code === 'DANGLING_REFERENCE')).toHaveLength(5);
  });

  it('reports invalid and cross-sheet ranges', () => {
    const fixture = validDocument();
    fixture.templates[0]!.printProfiles[0] = {
      ...fixture.templates[0]!.printProfiles[0]!,
      targets: [
        {
          type: 'range',
          range: {
            sheetId: 'sheet-2' as never,
            start: { row: 3, column: 0 },
            end: { row: 1, column: 2 },
          },
        },
      ],
    };

    expect(codesOf(fixture)).toContain('INVALID_RANGE');
  });

  it.each([
    {
      label: 'visibility',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.visibility = 'collapsed';
      },
    },
    {
      label: 'conditional range sheet',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'color-scale',
            range: {
              sheetId: 'missing-sheet',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            minimumColor: 'ff0000',
            maximumColor: '00ff00',
          },
        ];
      },
    },
    {
      label: 'conditional range order',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'color-scale',
            range: {
              sheetId: 'sheet-1',
              start: { row: 2, column: 0 },
              end: { row: 1, column: 1 },
            },
            minimumColor: 'ff0000',
            maximumColor: '00ff00',
          },
        ];
      },
    },
    {
      label: 'conditional operator',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'exec',
            formula: '1',
            style: {},
          },
        ];
      },
    },
    {
      label: 'conditional formula',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'between',
            formula: '=A1',
            style: {},
          },
        ];
      },
    },
    {
      label: 'conditional formula function',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'equal',
            formula: 'HYPERLINK(A1)',
            style: {},
          },
        ];
      },
    },
    {
      label: 'conditional formula sheet',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'equal',
            formula: 'Missing!A1',
            style: {},
          },
        ];
      },
    },
    {
      label: 'conditional style',
      mutate: (sheet: Record<string, unknown>) => {
        sheet.conditionalFormatting = [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'equal',
            formula: '1',
            style: { italic: true },
          },
        ];
      },
    },
  ])('rejects invalid worksheet $label semantics atomically', ({ mutate }) => {
    const fixture = validDocument();
    mutate(fixture.workbook.sheets[0]! as unknown as Record<string, unknown>);

    expect(codesOf(fixture)).toContain('DOCUMENT_SCHEMA_INVALID');
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

  it('preserves the extension diagnostic contract for non-JSON values', () => {
    const result = parseSpreadsheetDocument({
      ...validDocument(),
      extensions: { 'vendor.example': undefined },
    });

    expect(result.ok).toBe(false);
    if (result.ok) throw new Error('expected invalid document');
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'INVALID_EXTENSION_DATA',
        domain: 'extension',
        stage: 'decode',
        details: { path: '$.extensions.vendor.example' },
      }),
    );
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

  it('enforces configured sparse row and column layout limits', () => {
    const fixture = validDocument();
    Object.assign(fixture.workbook.sheets[0]!, {
      rows: [{ index: 1, height: 20 }],
      columns: [{ index: 1, width: 80 }],
    });

    expect(codesOf(fixture, { limits: { maxRows: 0 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
    expect(codesOf(fixture, { limits: { maxColumns: 0 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it.each([{ rows: [{ index: 0, height: -1 }] }, { columns: [{ index: 0, width: -1 }] }])(
    'rejects negative normalized sheet geometry: %o',
    (layout) => {
      const fixture = validDocument();
      Object.assign(fixture.workbook.sheets[0]!, layout);

      expect(codesOf(fixture)).toContain('DOCUMENT_SCHEMA_INVALID');
    },
  );

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

  it.each([
    ['maxSheets', Number.NaN],
    ['maxCells', Number.POSITIVE_INFINITY],
    ['maxMerges', -1],
    ['maxBytes', 1.5],
  ] as const)('rejects invalid %s limit values before decoding', (name, value) => {
    const input = Object.defineProperty(validDocument(), 'workbook', {
      get: () => {
        throw new Error('document decoding must not begin');
      },
    });

    const result = parseSpreadsheetDocument(input, {
      limits: { [name]: value },
    });

    expect(result).toMatchObject({
      ok: false,
      diagnostics: [
        {
          code: 'DOCUMENT_LIMIT_EXCEEDED',
          details: { path: `$.limits.${name}` },
        },
      ],
    });
  });

  it('treats undefined limits as absent instead of disabling defaults', () => {
    const sheets = Array.from({ length: 1_001 }, () => ({
      id: 'unused',
      name: 'Unused',
      cells: [],
      merges: [],
    }));
    Object.defineProperty(sheets, 0, {
      get: () => {
        throw new Error('over-limit sheet entries must not be decoded');
      },
    });
    const fixture = validDocument();
    fixture.workbook.sheets = sheets;

    expect(
      codesOf(fixture, {
        limits: { maxSheets: undefined },
      }),
    ).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('short-circuits sheet limits before decoding sheet entries', () => {
    const sheets = Array.from({ length: 2 }, () => ({
      id: 'unused',
      name: 'Unused',
      cells: [],
      merges: [],
    }));
    Object.defineProperty(sheets, 0, {
      get: () => {
        throw new Error('over-limit sheet entries must not be decoded');
      },
    });
    const fixture = validDocument();
    fixture.workbook.sheets = sheets;

    expect(codesOf(fixture, { limits: { maxSheets: 1 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('short-circuits cell limits before decoding cells or checking references', () => {
    const cells = Array.from({ length: 2 }, () => ({
      row: 0,
      column: 0,
      cell: { input: { type: 'blank' as const } },
    }));
    Object.defineProperty(cells, 0, {
      get: () => {
        throw new Error('over-limit cells must not be decoded');
      },
    });
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.cells = cells;

    expect(codesOf(fixture, { limits: { maxCells: 1 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('short-circuits merge limits before decoding merges or checking overlap', () => {
    const merges = Array.from({ length: 2 }, () => ({
      start: { row: 0, column: 0 },
      end: { row: 0, column: 0 },
    }));
    Object.defineProperty(merges, 0, {
      get: () => {
        throw new Error('over-limit merges must not be decoded');
      },
    });
    const fixture = validDocument();
    fixture.workbook.sheets[0]!.merges = merges;

    expect(codesOf(fixture, { limits: { maxMerges: 1 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('short-circuits object byte limits without stringifying the complete input', () => {
    const fixture = {
      schemaVersion: 2,
      id: 'document-byte-limit',
      workbook: {
        sheets: [],
        styles: [],
        validations: [],
        settings: { dateSystem: 'excel-1900' },
      },
      padding: 'x'.repeat(100),
    } as Record<string, unknown>;
    Object.defineProperty(fixture, 'extensions', {
      enumerable: true,
      get: () => {
        throw new Error('byte preflight must stop before later properties');
      },
    });

    expect(codesOf(fixture, { limits: { maxBytes: 16 } })).toContain('DOCUMENT_LIMIT_EXCEEDED');
  });

  it('reports hostile object accessors atomically during shallow preflight', () => {
    const fixture = Object.defineProperty({ schemaVersion: 2 }, 'workbook', {
      enumerable: true,
      get: () => {
        throw new Error('hostile accessor');
      },
    });

    expect(() => parseSpreadsheetDocument(fixture)).not.toThrow();
    expect(codesOf(fixture)).toContain('DOCUMENT_SCHEMA_INVALID');
  });

  it('rejects a stateful workbook getter instead of letting it bypass maxSheets', () => {
    const decodedWorkbook = validDocument().workbook;
    decodedWorkbook.sheets.push({
      id: 'sheet-2',
      name: 'Sheet 2',
      cells: [],
      merges: [],
    });

    expect(
      codesOf(documentWithWorkbookGetter(decodedWorkbook), { limits: { maxSheets: 1 } }),
    ).toContain('DOCUMENT_SCHEMA_INVALID');
  });

  it('rejects a stateful workbook getter instead of letting it bypass maxCells', () => {
    const decodedWorkbook = validDocument().workbook;
    decodedWorkbook.sheets[0]!.cells.push(
      { row: 0, column: 0, cell: { input: { type: 'blank' } } },
      { row: 1, column: 0, cell: { input: { type: 'blank' } } },
    );

    expect(
      codesOf(documentWithWorkbookGetter(decodedWorkbook), { limits: { maxCells: 1 } }),
    ).toContain('DOCUMENT_SCHEMA_INVALID');
  });

  it('rejects a stateful workbook getter instead of letting it bypass maxMerges', () => {
    const decodedWorkbook = validDocument().workbook;
    decodedWorkbook.sheets[0]!.merges.push(
      { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
      { start: { row: 1, column: 0 }, end: { row: 1, column: 0 } },
    );

    expect(
      codesOf(documentWithWorkbookGetter(decodedWorkbook), { limits: { maxMerges: 1 } }),
    ).toContain('DOCUMENT_SCHEMA_INVALID');
  });

  it('rejects a stateful workbook getter instead of letting it bypass maxBytes', () => {
    const decodedWorkbook = validDocument().workbook;
    decodedWorkbook.sheets[0]!.cells.push({
      row: 0,
      column: 0,
      cell: { input: { type: 'string', value: 'x'.repeat(10_000) } },
    });

    expect(
      codesOf(documentWithWorkbookGetter(decodedWorkbook), { limits: { maxBytes: 1_000 } }),
    ).toContain('DOCUMENT_SCHEMA_INVALID');
  });

  it.each([
    ['bigint', 1n],
    ['undefined', undefined],
    ['function', () => undefined],
    ['symbol', Symbol('not-json')],
    ['NaN', Number.NaN],
    ['Infinity', Number.POSITIVE_INFINITY],
  ])('rejects non-JSON %s values even in unknown fields', (_label, value) => {
    const fixture = validDocument() as unknown as Record<string, unknown>;
    fixture.unknown = value;

    expect(() => parseSpreadsheetDocument(fixture)).not.toThrow();
    expect(codesOf(fixture)).toContain('DOCUMENT_SCHEMA_INVALID');
  });
});
