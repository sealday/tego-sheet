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

  it('reports unsupported known and unknown fields without copying them', () => {
    const { document, diagnostics } = migrated(fixture('complete'));
    const degradedPaths = diagnostics
      .filter(({ code }) => code === 'LEGACY_FIELD_DEGRADED')
      .map(({ details }) => (details as { path?: string } | undefined)?.path);

    expect(degradedPaths).toEqual(
      expect.arrayContaining([
        '$[0].freeze',
        '$[0].rows.len',
        '$[0].rows.0.height',
        '$[0].rows.0.hide',
        '$[0].cols.len',
        '$[0].cols.0.width',
        '$[0].cols.0.hide',
        '$[0].autofilter',
        '$[0].rows.0.cells.0.editable',
        '$[0].rows.0.cells.0.printable',
      ]),
    );
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
      { rows: { bad: { cells: {} } } },
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
});
