import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  migrateLegacyWorkbook,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
} from '../../../src/document';
import type { MigrationResult, SpreadsheetDocument } from '../../../src/document';

const fixture = (name: string): unknown =>
  JSON.parse(
    readFileSync(
      fileURLToPath(new URL(`../../fixtures/document/legacy/${name}.json`, import.meta.url)),
      'utf8',
    ),
  );

const deterministicIds = {
  documentId: () => 'document-migrated',
  sheetId: (index: number) => `sheet-${index + 1}`,
} as const;

function migrated(input: unknown): Extract<MigrationResult, { readonly ok: true }> {
  const result = migrateLegacyWorkbook(input, { ids: deterministicIds });
  expect(result.ok, JSON.stringify(result.diagnostics)).toBe(true);
  if (!result.ok) throw new Error('expected migration success');
  return result;
}

function cellAt(document: SpreadsheetDocument, row: number, column: number) {
  return document.workbook.sheets[0]?.cells.find(
    (entry) => entry.row === row && entry.column === column,
  )?.cell;
}

describe('legacy WorkbookInput migration', () => {
  it.each([
    'autofilter',
    'blank-object',
    'cells',
    'columns',
    'empty-array',
    'multiple-sheets',
    'rows',
    'sheet-fields',
    'sparse-falsy',
    'styles',
    'validations',
  ])('migrates the current %s serialization fixture', (name) => {
    const current = (
      fixture(`../../../parity/fixtures/workbooks/${name}`) as {
        input: unknown;
      }
    ).input;

    const result = migrateLegacyWorkbook(current, { ids: deterministicIds });
    expect(result.ok, JSON.stringify(result.diagnostics)).toBe(true);
  });

  it('persists legacy dimensions, visibility, defaults, freeze, filter, and print flags', () => {
    const { document, diagnostics } = migrated(fixture('complete'));
    const sheet = document.workbook.sheets[0]!;

    expect(sheet).toMatchObject({
      rowCount: 10001,
      columnCount: 53,
      freeze: { row: 2, column: 2 },
      rows: [{ index: 0, height: 28, hidden: false, styleId: 'legacy-style-0' }],
      columns: [{ index: 0, width: 144, hidden: true, styleId: 'legacy-style-0' }],
      filter: {
        range: { start: { row: 0, column: 0 }, end: { row: 9, column: 4 } },
        filters: [{ column: 0, operator: 'in', values: ['kept'] }],
        sort: { column: 0, direction: 'asc' },
      },
    });
    expect(cellAt(document, 0, 0)).toMatchObject({ editable: false, printable: false });
    expect(diagnostics.map(({ code }) => code)).not.toContain('LEGACY_FIELD_DEGRADED');
  });

  it('migrates single-sheet and ordered multi-sheet inputs with injected stable IDs', () => {
    const single = migrated(fixture('complete'));
    const multiple = migrated(fixture('multiple-sheets'));

    expect(single.document).toMatchObject({
      schemaVersion: 2,
      id: 'document-migrated',
      workbook: { sheets: [{ id: 'sheet-1', name: 'Legacy report' }] },
    });
    expect(multiple.document.workbook.sheets.map(({ id, name }) => ({ id, name }))).toEqual([
      { id: 'sheet-1', name: 'First' },
      { id: 'sheet-2', name: 'Second' },
    ]);
  });

  it('preserves formula source, explicit falsy inputs, blank cells, and sparse high indexes', () => {
    const { document } = migrated(fixture('complete'));

    expect(cellAt(document, 0, 0)?.input).toEqual({ type: 'formula', source: '=1+1' });
    expect(cellAt(document, 0, 1)?.input).toEqual({ type: 'string', value: '' });
    expect(cellAt(document, 0, 2)?.input).toEqual({ type: 'number', value: 0 });
    expect(cellAt(document, 0, 3)?.input).toEqual({ type: 'boolean', value: false });
    expect(cellAt(document, 0, 4)?.input).toEqual({ type: 'blank' });
    expect(cellAt(document, 10_000, 52)?.input).toEqual({ type: 'string', value: 'sparse' });
  });

  it('deduplicates equivalent styles and never treats a cached formula value as truth', () => {
    const { document } = migrated(fixture('complete'));

    expect(document.workbook.styles).toHaveLength(1);
    expect(cellAt(document, 0, 0)).toMatchObject({
      input: { type: 'formula', source: '=1+1' },
      styleId: document.workbook.styles[0]?.id,
    });
    expect(cellAt(document, 0, 1)?.styleId).toBe(document.workbook.styles[0]?.id);
  });

  it('normalizes merges and maps reliable validation references onto explicit cells', () => {
    const { document } = migrated(fixture('complete'));

    expect(document.workbook.sheets[0]?.merges).toEqual([
      { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
    ]);
    expect(document.workbook.validations).toHaveLength(1);
    expect(cellAt(document, 0, 0)?.validationId).toBe(document.workbook.validations[0]?.id);
    expect(cellAt(document, 1, 0)).toMatchObject({
      input: { type: 'blank' },
      validationId: document.workbook.validations[0]?.id,
    });
  });

  it('reports unknown fields without copying them', () => {
    const { document, diagnostics } = migrated(fixture('complete'));
    expect(diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'LEGACY_FIELD_DROPPED',
        details: expect.objectContaining({ path: '$[0].vendorSheet' }),
      }),
    );
    expect(diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'LEGACY_FIELD_DROPPED',
        details: expect.objectContaining({ path: '$[0].styles[0].font.vendorFont' }),
      }),
    );
    expect(document.workbook.styles[0]?.value).not.toHaveProperty('font.vendorFont');
    expect(document.extensions).toEqual({});
  });

  it('does not mutate input or share mutable references with the returned frozen document', () => {
    const input = fixture('complete');
    const before = structuredClone(input);
    const { document } = migrated(input);

    expect(input).toEqual(before);
    expect(Object.isFrozen(document)).toBe(true);
    expect(document.workbook.styles[0]?.value).not.toBe(
      (input as { styles: readonly unknown[] }).styles[0],
    );
  });

  it('is deterministic with injected IDs and produces stable schema 2 bytes', () => {
    const input = fixture('complete');
    const first = migrated(input);
    const second = migrated(structuredClone(input));
    const firstBytes = serializeSpreadsheetDocument(first.document);

    expect(serializeSpreadsheetDocument(second.document)).toBe(firstBytes);
    const parsed = parseSpreadsheetDocument(firstBytes);
    expect(parsed.ok).toBe(true);
    if (parsed.ok) expect(serializeSpreadsheetDocument(parsed.document)).toBe(firstBytes);
  });

  it('returns atomic located errors for invalid legacy values and references', () => {
    const invalidCases = [
      { rows: { 0: { cells: { 0: { type: 'number', value: Number.NaN } } } } },
      { styles: [], rows: { 0: { cells: { 0: { text: 'x', style: 3 } } } } },
      { merges: ['not-a-range'] },
    ];

    for (const input of invalidCases) {
      const result = migrateLegacyWorkbook(input, { ids: deterministicIds });
      expect(result.ok).toBe(false);
      expect('document' in result).toBe(false);
      expect(result.diagnostics).toEqual([
        expect.objectContaining({
          domain: 'document',
          stage: 'migrate',
          severity: 'error',
          details: expect.objectContaining({ path: expect.any(String) }),
        }),
      ]);
    }
  });

  it('rejects JavaScript Date values rather than persisting Date objects', () => {
    const result = migrateLegacyWorkbook(
      { rows: { 0: { cells: { 0: { value: new Date('2026-01-02T00:00:00Z') } } } } },
      { ids: deterministicIds },
    );

    expect(result).toMatchObject({
      ok: false,
      diagnostics: [
        {
          code: 'LEGACY_VALUE_INVALID',
          details: { path: '$[0].rows.0.cells.0.value' },
        },
      ],
    });
  });

  it('ignores a cached Date when authoritative text supplies the input', () => {
    const cached = new Date('2026-01-02T00:00:00Z');
    const result = migrated({ rows: { 0: { cells: { 0: { text: '=TODAY()', value: cached } } } } });

    expect(cellAt(result.document, 0, 0)?.input).toEqual({
      type: 'formula',
      source: '=TODAY()',
    });
  });

  it('deduplicates identical sheet and cell merge representations', () => {
    const current = (
      fixture('../../../parity/fixtures/workbooks/sheet-fields') as {
        input: unknown;
      }
    ).input;
    const result = migrated(current);

    expect(result.document.workbook.sheets[0]?.merges).toEqual([
      { start: { row: 0, column: 0 }, end: { row: 1, column: 1 } },
    ]);
  });

  it('normalizes leading-zero sparse keys and drops collection extensions', () => {
    const result = migrated({
      rows: {
        '01': { cells: { '002': { text: 'normalized' }, vendorCells: { ignored: true } } },
        vendorRows: { ignored: true },
      },
      cols: { '003': { width: 80 }, vendorCols: false },
    });

    expect(cellAt(result.document, 1, 2)?.input).toEqual({
      type: 'string',
      value: 'normalized',
    });
    expect(result.document.workbook.sheets[0]?.columns).toEqual([{ index: 3, width: 80 }]);
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ details: { path: '$[0].rows.vendorRows' } }),
        expect.objectContaining({ details: { path: '$[0].rows.01.cells.vendorCells' } }),
        expect.objectContaining({ details: { path: '$[0].cols.vendorCols' } }),
      ]),
    );
  });

  it.each([
    { styles: [{ font: { bold: 'yes' } }] },
    { styles: [{ align: 'diagonal' }] },
    { validations: [{ refs: ['A1'], mode: 'cell', type: 'bogus', required: false }] },
    { validations: [{ refs: ['A1'], mode: 'cell', type: 'number', required: 'no' }] },
    {
      validations: [
        { refs: ['A1'], mode: 'cell', type: 'number', required: false, operator: 'bogus' },
      ],
    },
    {
      validations: [
        {
          refs: ['A1'],
          mode: 'cell',
          type: 'list',
          required: false,
          operator: 'in',
          value: 'not-a-list',
        },
      ],
    },
    { rows: { 0: { cells: { 0: { type: 'bogus' } } } } },
  ])('rejects invalid known legacy semantics atomically: %o', (input) => {
    const result = migrateLegacyWorkbook(input, { ids: deterministicIds });

    expect(result).toMatchObject({
      ok: false,
      diagnostics: [expect.objectContaining({ code: 'LEGACY_VALUE_INVALID' })],
    });
  });

  it('rejects sparse key collisions after leading-zero normalization', () => {
    const result = migrateLegacyWorkbook(
      { rows: { 1: { cells: {} }, '01': { cells: {} } } },
      { ids: deterministicIds },
    );

    expect(result).toMatchObject({
      ok: false,
      diagnostics: [
        expect.objectContaining({
          code: 'LEGACY_VALUE_INVALID',
          details: { path: '$[0].rows.01' },
        }),
      ],
    });
  });

  it('enforces the cumulative validation expansion limit', () => {
    const result = migrateLegacyWorkbook(
      {
        validations: [
          {
            refs: ['A1:A600000', 'B1:B600000'],
            mode: 'cell',
            type: 'number',
            required: false,
          },
        ],
      },
      { ids: deterministicIds },
    );

    expect(result).toMatchObject({
      ok: false,
      diagnostics: [
        expect.objectContaining({
          code: 'LEGACY_REFERENCE_INVALID',
          details: { path: '$[0].validations[0].refs[1]' },
        }),
      ],
    });
  });

  it('returns an atomic failure instead of invoking accessor properties', () => {
    const input = Object.defineProperty({}, 'name', {
      enumerable: true,
      get: () => {
        throw new Error('boom');
      },
    });

    expect(() => migrateLegacyWorkbook(input, { ids: deterministicIds })).not.toThrow();
    expect(migrateLegacyWorkbook(input, { ids: deterministicIds })).toMatchObject({
      ok: false,
      diagnostics: [
        expect.objectContaining({
          code: 'LEGACY_VALUE_INVALID',
          details: { path: '$[0].name' },
        }),
      ],
    });
  });
});
