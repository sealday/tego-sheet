import { describe, expect, it } from 'vitest';
import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
} from '../../../src/document';
import type {
  CellInput,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../../src/document';

const parseOk = (input: unknown): SpreadsheetDocument => {
  const result = parseSpreadsheetDocument(input);
  expect(result.ok, JSON.stringify(result.diagnostics)).toBe(true);
  if (!result.ok) throw new Error('expected a valid document');
  return result.document;
};

const documentWithInputs = (
  inputs: readonly CellInput[],
  dateSystem: 'excel-1900' | 'excel-1904' = 'excel-1900',
): SpreadsheetDocumentInput => ({
  schemaVersion: 2,
  id: 'document-roundtrip',
  workbook: {
    sheets: [
      {
        id: 'sheet-stable',
        name: 'Original name',
        cells: inputs.map((input, column) => ({
          row: 0,
          column,
          cell: { input },
        })),
        merges: [],
      },
    ],
    styles: [],
    validations: [],
    settings: { dateSystem, localeHint: 'en-US' },
  },
  templates: [],
  resources: { items: [] },
  extensions: {},
});

describe('Workbook 2.0 parsing and serialization', () => {
  it('round-trips every typed cell input including explicit falsy and custom JSON values', () => {
    const inputs: readonly CellInput[] = [
      { type: 'blank' },
      { type: 'string', value: '' },
      { type: 'number', value: 0 },
      { type: 'boolean', value: false },
      { type: 'formula', source: '=A1+1' },
      {
        type: 'custom',
        cellType: 'example.status',
        schemaVersion: 1,
        value: { nested: [null, false, 0, '', { z: 1, a: 2 }] },
      },
    ];

    const encoded = serializeSpreadsheetDocument(parseOk(documentWithInputs(inputs)));
    const reparsed = parseOk(encoded);

    expect(serializeSpreadsheetDocument(reparsed)).toBe(encoded);
    expect(reparsed.workbook.sheets[0]?.cells.map((entry) => entry.cell.input)).toEqual(inputs);
  });

  it('preserves the stable sheet ID when a copied fixture is renamed', () => {
    const fixture = documentWithInputs([{ type: 'string', value: 'kept' }]);
    const renamed = structuredClone(fixture);
    renamed.workbook.sheets[0]!.name = 'Renamed display label';

    const document = parseOk(renamed);

    expect(document.workbook.sheets[0]).toMatchObject({
      id: 'sheet-stable',
      name: 'Renamed display label',
    });
  });

  it('round-trips normalized sheet layout, filter, and print semantics', () => {
    const fixture = documentWithInputs([{ type: 'string', value: 'kept' }]);
    fixture.workbook.styles = [{ id: 'style-layout', value: { align: 'right' } }];
    Object.assign(fixture.workbook.sheets[0]!, {
      rowCount: 100,
      columnCount: 26,
      rows: [{ index: 2, height: 42, hidden: false, styleId: 'style-layout' }],
      columns: [{ index: 4, width: 72, hidden: true, styleId: 'style-layout' }],
      freeze: { row: 2, column: 2 },
      filter: {
        range: { start: { row: 0, column: 0 }, end: { row: 4, column: 2 } },
        filters: [{ column: 1, operator: 'in', values: ['open', ''] }],
        sort: { column: 2, direction: 'desc' },
      },
    });
    Object.assign(fixture.workbook.sheets[0]!.cells[0]!.cell, {
      editable: false,
      printable: false,
    });

    const document = parseOk(fixture);
    const reparsed = parseOk(serializeSpreadsheetDocument(document));

    expect(reparsed.workbook.sheets[0]).toMatchObject({
      rowCount: 100,
      columnCount: 26,
      rows: [{ index: 2, height: 42, hidden: false, styleId: 'style-layout' }],
      columns: [{ index: 4, width: 72, hidden: true, styleId: 'style-layout' }],
      freeze: { row: 2, column: 2 },
      filter: {
        range: { start: { row: 0, column: 0 }, end: { row: 4, column: 2 } },
        filters: [{ column: 1, operator: 'in', values: ['open', ''] }],
        sort: { column: 2, direction: 'desc' },
      },
      cells: [{ cell: expect.objectContaining({ editable: false, printable: false }) }],
    });
  });

  it.each(['excel-1900', 'excel-1904'] as const)(
    'preserves Excel date serial numbers with the explicit %s date system',
    (dateSystem) => {
      const serial = 45_205;
      const document = parseOk(documentWithInputs([{ type: 'number', value: serial }], dateSystem));

      expect(document.workbook.settings.dateSystem).toBe(dateSystem);
      expect(document.workbook.sheets[0]?.cells[0]?.cell.input).toEqual({
        type: 'number',
        value: serial,
      });
    },
  );

  it('creates a valid deeply frozen default document', () => {
    const document = createSpreadsheetDocument({
      id: 'new-document',
      sheetId: 'first-sheet',
      sheetName: 'Sheet 1',
      dateSystem: 'excel-1904',
    });

    expect(document).toMatchObject({
      schemaVersion: 2,
      id: 'new-document',
      workbook: {
        settings: { dateSystem: 'excel-1904' },
        sheets: [{ id: 'first-sheet', name: 'Sheet 1' }],
      },
    });
    expect(Object.isFrozen(document)).toBe(true);
    expect(Object.isFrozen(document.workbook.sheets)).toBe(true);
  });

  it('creates collision-resistant opaque IDs when callers do not provide them', () => {
    const first = createSpreadsheetDocument();
    const second = createSpreadsheetDocument();
    const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

    expect(first.id).toMatch(uuidPattern);
    expect(first.workbook.sheets[0]?.id).toMatch(uuidPattern);
    expect(new Set([first.id, first.workbook.sheets[0]?.id, second.id])).toHaveLength(3);
  });

  it('serializes registries, sparse coordinates, extension keys, and JSON keys canonically', () => {
    const fixture = documentWithInputs([]);
    fixture.workbook.styles = [
      { id: 'style-z', value: { z: 1, a: 2 } },
      { id: 'style-a', value: null },
    ];
    fixture.workbook.sheets[0]!.cells = [
      { row: 10, column: 0, cell: { input: { type: 'blank' } } },
      { row: 2, column: 5, cell: { input: { type: 'string', value: 'later column' } } },
      { row: 2, column: 1, cell: { input: { type: 'number', value: 1 } } },
    ];
    fixture.extensions = {
      'vendor.z': { z: true, a: false },
      'vendor.a': null,
    };

    const encoded = serializeSpreadsheetDocument(parseOk(fixture));
    const plain = JSON.parse(encoded) as {
      workbook: {
        styles: { id: string }[];
        sheets: { cells: { row: number; column: number }[] }[];
      };
      extensions: Record<string, unknown>;
    };

    expect(plain.workbook.styles.map(({ id }) => id)).toEqual(['style-a', 'style-z']);
    expect(plain.workbook.sheets[0]?.cells.map(({ row, column }) => [row, column])).toEqual([
      [2, 1],
      [2, 5],
      [10, 0],
    ]);
    expect(Object.keys(plain.extensions)).toEqual(['vendor.a', 'vendor.z']);
    expect(encoded.indexOf('"a":2')).toBeLessThan(encoded.indexOf('"z":1'));
  });

  it('uses locale-independent code-unit ordering for opaque registry IDs', () => {
    const fixture = documentWithInputs([]);
    fixture.workbook.styles = ['é', 'z', 'ä'].map((id) => ({ id, value: null }));

    const document = parseOk(fixture);

    expect(document.workbook.styles.map(({ id }) => id)).toEqual(['z', 'ä', 'é']);
  });

  it('produces identical bytes for semantically equal unordered registries and resources', () => {
    const first = documentWithInputs([]);
    first.workbook.styles = [
      { id: 'style-z', value: { z: 1, a: 2 } },
      { id: 'style-a', value: null },
    ];
    first.workbook.validations = [
      { id: 'validation-z', value: { z: true, a: false } },
      { id: 'validation-a', value: null },
    ];
    first.resources.items = [
      { id: 'resource-z', kind: 'image', metadata: { z: 1, a: 2 } },
      { id: 'resource-a', kind: 'font' },
    ];
    const second = structuredClone(first);
    second.workbook.styles.reverse();
    second.workbook.validations.reverse();
    second.resources.items.reverse();

    expect(serializeSpreadsheetDocument(parseOk(first))).toBe(
      serializeSpreadsheetDocument(parseOk(second)),
    );
  });

  it('preserves user-defined sheet and template order while canonicalizing bytes', () => {
    const fixture = documentWithInputs([]);
    fixture.workbook.sheets.push({
      id: 'sheet-first-lexically',
      name: 'Second user sheet',
      cells: [],
      merges: [],
    });
    fixture.templates = [
      {
        id: 'template-z',
        name: 'First template',
        sheetId: 'sheet-stable',
        printProfile: {
          paperSize: 'A4',
          orientation: 'portrait',
          margins: { top: 0, right: 0, bottom: 0, left: 0 },
        },
      },
      {
        id: 'template-a',
        name: 'Second template',
        sheetId: 'sheet-stable',
        printProfile: {
          paperSize: 'A4',
          orientation: 'portrait',
          margins: { top: 0, right: 0, bottom: 0, left: 0 },
        },
      },
    ];

    const document = parseOk(fixture);

    expect(document.workbook.sheets.map(({ id }) => id)).toEqual([
      'sheet-stable',
      'sheet-first-lexically',
    ]);
    expect(document.templates.map(({ id }) => id)).toEqual(['template-z', 'template-a']);
  });

  it('preserves JSON meta-property keys without prototype mutation', () => {
    const fixture = documentWithInputs([]);
    fixture.extensions['vendor.meta'] = JSON.parse(
      '{"constructor":{"safe":true},"__proto__":{"polluted":false}}',
    ) as never;

    const encoded = serializeSpreadsheetDocument(parseOk(fixture));
    const document = parseOk(encoded);
    const value = document.extensions['vendor.meta'] as Record<string, unknown>;

    expect(Object.getPrototypeOf(value)).toBeNull();
    expect(Object.keys(value)).toEqual(['__proto__', 'constructor']);
    expect(value.__proto__).toEqual({ polluted: false });
    expect(serializeSpreadsheetDocument(document)).toBe(encoded);
  });

  it('isolates caller input and deeply freezes the returned snapshot', () => {
    const fixture = documentWithInputs([
      {
        type: 'custom',
        cellType: 'example.payload',
        schemaVersion: 1,
        value: { nested: { label: 'original' } },
      },
    ]);
    const document = parseOk(fixture);
    const encoded = serializeSpreadsheetDocument(document);

    const originalValue = fixture.workbook.sheets[0]!.cells[0]!.cell.input;
    if (originalValue.type !== 'custom') throw new Error('expected custom input');
    (originalValue.value as { nested: { label: string } }).nested.label = 'changed input';

    const snapshotValue = document.workbook.sheets[0]!.cells[0]!.cell.input;
    if (snapshotValue.type !== 'custom') throw new Error('expected custom snapshot input');
    expect(() => {
      (snapshotValue.value as { nested: { label: string } }).nested.label = 'changed snapshot';
    }).toThrow(TypeError);

    expect(serializeSpreadsheetDocument(document)).toBe(encoded);
    expect(snapshotValue.value).toEqual({ nested: { label: 'original' } });
  });
});
