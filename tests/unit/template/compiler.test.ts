import { describe, expect, it } from 'vitest';
import type { SpreadsheetDocument } from '../../../src/document';
import { compileSpreadsheetTemplate } from '../../../src/template/compiler';
import { hashSpreadsheetDocument } from '../../../src/template/hash';
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
  it('resolves the canonical persisted template by ID and hashes template/profile changes', () => {
    const sourceTemplate = template();
    const persisted = { ...document, templates: [sourceTemplate] } as SpreadsheetDocument;
    const compiled = compileSpreadsheetTemplate(persisted, sourceTemplate.id);
    expect(compiled.hasErrors).toBe(false);
    expect(compiled.template?.sourceDocumentHash).toBe(hashSpreadsheetDocument(persisted));

    const changed = {
      ...persisted,
      templates: [
        {
          ...sourceTemplate,
          printProfiles: [{ ...sourceTemplate.printProfiles[0]!, showGridlines: false }],
        },
      ],
    } as SpreadsheetDocument;
    expect(hashSpreadsheetDocument(changed)).not.toBe(compiled.template?.sourceDocumentHash);
  });

  it('isolates the compiled snapshot without freezing caller-owned inputs', () => {
    const sourceTemplate = template();
    const persisted = { ...document, templates: [sourceTemplate] } as SpreadsheetDocument;
    const result = compileSpreadsheetTemplate(persisted, sourceTemplate.id);
    expect(result.hasErrors).toBe(false);
    expect(Object.isFrozen(persisted)).toBe(false);
    expect(Object.isFrozen(sourceTemplate)).toBe(false);
    expect(result.template?.sourceDocument).not.toBe(persisted);
    expect(result.template?.ir.template).not.toBe(sourceTemplate);
  });

  it('aggregates invalid IDs, binding ranges, breaks, paper and scale geometry', () => {
    const invalid = template({
      bindings: [
        {
          id: 'duplicate' as never,
          type: 'value',
          target: { sheetId: 'missing' as never, row: 0, column: 0 },
          expression: 'value',
        },
        {
          id: 'duplicate' as never,
          type: 'repeat-rows',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 3, column: 1 },
            end: { row: 1, column: 0 },
          },
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
      ],
      printProfiles: [
        {
          ...template().printProfiles[0]!,
          id: 'duplicate-profile',
          page: {
            ...template().printProfiles[0]!.page,
            paper: { type: 'custom', width: 0, height: Number.NaN },
            scale: { type: 'fit-width', pages: 0 },
          },
          manualBreaks: [{ sheetId: 'missing' as never, beforeRow: -1 }],
          repeatRows: {
            sheetId: 'sheet-1' as never,
            start: { row: 2, column: 0 },
            end: { row: 1, column: 0 },
          },
        },
        { ...template().printProfiles[0]!, id: 'duplicate-profile' },
      ],
    });
    const persisted = { ...document, templates: [invalid] } as SpreadsheetDocument;
    const result = compileSpreadsheetTemplate(persisted, invalid.id);
    expect(result.hasErrors).toBe(true);
    expect(result.diagnostics.map(({ code }) => code)).toEqual(
      expect.arrayContaining([
        'DUPLICATE_BINDING_ID',
        'DUPLICATE_PRINT_PROFILE_ID',
        'INVALID_BINDING_TARGET',
        'INVALID_BINDING_RANGE',
        'INVALID_PAGE_BREAK',
        'INVALID_PAGE_GEOMETRY',
        'INVALID_PRINT_SCALE',
        'INVALID_REPEAT_TITLE_RANGE',
      ]),
    );
  });

  it('returns an atomic diagnostic for malformed JavaScript input', () => {
    const malformed = {
      ...template(),
      bindings: [
        {
          id: 'broken',
          type: 'repeat-rows',
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const persisted = { ...document, templates: [malformed] } as SpreadsheetDocument;
    expect(() => compileSpreadsheetTemplate(persisted, malformed.id)).not.toThrow();
    expect(compileSpreadsheetTemplate(persisted, malformed.id)).toEqual(
      expect.objectContaining({
        hasErrors: true,
        diagnostics: expect.arrayContaining([
          expect.objectContaining({ code: 'INVALID_TEMPLATE_STRUCTURE' }),
        ]),
      }),
    );
  });

  it('rejects malformed nested print profile structures without throwing', () => {
    const malformed = {
      ...template(),
      printProfiles: [
        {
          id: 'bad-profile',
          name: 'Bad',
          targets: [],
          page: {},
          manualBreaks: [],
          showGridlines: true,
          showHeadings: false,
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const persisted = { ...document, templates: [malformed] } as SpreadsheetDocument;
    expect(() => compileSpreadsheetTemplate(persisted, malformed.id)).not.toThrow();
    expect(compileSpreadsheetTemplate(persisted, malformed.id)).toEqual(
      expect.objectContaining({
        hasErrors: true,
        diagnostics: [expect.objectContaining({ code: 'INVALID_TEMPLATE_STRUCTURE' })],
      }),
    );
  });

  it('rejects oversized compiler input before cloning the source document', () => {
    const binding = {
      id: 'oversized' as never,
      type: 'value' as const,
      target: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
      expression: 'value',
    };
    const oversized = {
      ...template(),
      bindings: Array.from({ length: 10_001 }, () => binding),
    };
    const persisted = { ...document, templates: [oversized] } as SpreadsheetDocument;
    expect(compileSpreadsheetTemplate(persisted, oversized.id)).toEqual(
      expect.objectContaining({
        hasErrors: true,
        diagnostics: [expect.objectContaining({ code: 'COMPILATION_RESOURCE_LIMIT' })],
      }),
    );
  });

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
