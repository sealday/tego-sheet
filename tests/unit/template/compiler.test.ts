import { describe, expect, it } from 'vitest';
import type { SpreadsheetDocument } from '../../../src/document';
import { compileSpreadsheetTemplate } from '../../../src/template/compiler';
import type { SpreadsheetTemplate } from '../../../src/template/model';

const document = {
  schemaVersion: 2,
  id: 'document-1',
  workbook: {
    sheets: [
      {
        id: 'sheet-1',
        name: 'Invoice',
        cells: [],
        merges: [{ start: { row: 0, column: 0 }, end: { row: 2, column: 0 } }],
        rows: [],
        columns: [],
      },
    ],
    styles: [],
    validations: [],
    settings: { dateSystem: 'excel-1900' },
  },
  templates: [],
  resources: { items: [] },
  extensions: {},
} as unknown as SpreadsheetDocument;

function template(overrides: Partial<SpreadsheetTemplate> = {}): SpreadsheetTemplate {
  return {
    id: 'template-1' as never,
    name: 'Invoice',
    bindings: [],
    printProfiles: [
      {
        id: 'profile-1',
        name: 'A4',
        targets: [{ type: 'sheet', sheetId: 'sheet-1' as never }],
        page: {
          paper: { type: 'A4' },
          orientation: 'portrait',
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
          scale: { type: 'fixed', value: 1 },
        },
        manualBreaks: [],
        showGridlines: true,
        showHeadings: false,
      },
    ],
    ...overrides,
  };
}

describe('template compiler', () => {
  it('compiles metadata bindings into immutable IR and hashes the complete source', () => {
    const result = compileSpreadsheetTemplate(document, template());
    expect(result.hasErrors).toBe(false);
    expect(result.template).toMatchObject({
      templateId: 'template-1',
      compilerVersion: expect.any(String),
      sourceDocumentHash: expect.stringMatching(/^[a-f0-9]{16}$/u),
    });
    expect(Object.isFrozen(result.template?.ir)).toBe(true);
  });

  it('aggregates expression, overlap, target, and merge-boundary diagnostics', () => {
    const result = compileSpreadsheetTemplate(
      document,
      template({
        bindings: [
          {
            id: 'repeat-a' as never,
            type: 'repeat-rows',
            range: {
              sheetId: 'sheet-1' as never,
              start: { row: 0, column: 0 },
              end: { row: 1, column: 1 },
            },
            source: 'items',
            empty: 'remove',
            pageBreak: 'auto',
          },
          {
            id: 'conditional-a' as never,
            type: 'conditional-range',
            range: {
              sheetId: 'sheet-1' as never,
              start: { row: 1, column: 0 },
              end: { row: 2, column: 1 },
            },
            when: 'globalThis.enabled',
          },
        ],
        printProfiles: [
          {
            ...template().printProfiles[0]!,
            targets: [{ type: 'sheet', sheetId: 'missing' as never }],
          },
        ],
      }),
    );
    expect(result.hasErrors).toBe(true);
    expect(result.diagnostics.map(({ code }) => code)).toEqual(
      expect.arrayContaining([
        'OVERLAPPING_REPEAT_REGION',
        'MERGE_CROSSES_REPEAT_BOUNDARY',
        'INVALID_EXPRESSION',
        'INVALID_PRINT_TARGET',
      ]),
    );
    expect(result.template).toBeUndefined();
  });
});
