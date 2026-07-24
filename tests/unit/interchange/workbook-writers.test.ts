import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { parseSpreadsheetDocument, type SpreadsheetDocument } from '../../../src/document';
import {
  createOdsReader,
  createOdsWriter,
  createXlsxReader,
  createXlsxWriter,
  type WorkbookExportResult,
  type WorkbookWriter,
} from '../../../src/interchange';

function semanticDocument(): SpreadsheetDocument {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'semantic-round-trip',
    workbook: {
      sheets: [
        {
          id: 'sheet-data',
          name: 'Data',
          cells: [
            {
              row: 0,
              column: 0,
              cell: {
                input: { type: 'string', value: 'Heading' },
                styleId: 'heading',
                validationId: 'required-list',
              },
            },
            { row: 1, column: 0, cell: { input: { type: 'number', value: 42.5 } } },
            { row: 1, column: 1, cell: { input: { type: 'boolean', value: true } } },
            { row: 2, column: 0, cell: { input: { type: 'formula', source: '=A2+1' } } },
          ],
          merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 1 } }],
          filter: {
            range: { start: { row: 0, column: 0 }, end: { row: 2, column: 1 } },
            filters: [{ column: 0, operator: 'in', values: ['Heading', '42.5'] }],
            sort: { column: 0, direction: 'asc' },
          },
          conditionalFormatting: [
            {
              type: 'cell-is',
              range: {
                sheetId: 'sheet-data',
                start: { row: 1, column: 0 },
                end: { row: 2, column: 0 },
              },
              operator: 'greaterThan',
              formula: '40',
              style: { color: '#ffffff', backgroundColor: '#008800', bold: true },
            },
          ],
        },
      ],
      styles: [
        {
          id: 'heading',
          value: {
            color: '#112233',
            backgroundColor: '#ddeeff',
            bold: true,
            horizontalAlign: 'center',
            verticalAlign: 'middle',
            numberFormat: '0.00',
          },
        },
      ],
      validations: [
        {
          id: 'required-list',
          value: {
            type: 'list',
            formula1: '"Heading,42.5"',
            allowBlank: false,
          },
        },
      ],
      settings: { dateSystem: 'excel-1900', localeHint: 'en-US' },
    },
    templates: [
      {
        id: 'template-print',
        name: 'Print',
        bindings: [],
        printProfiles: [
          {
            id: 'profile-main',
            name: 'Main',
            targets: [{ type: 'sheet', sheetId: 'sheet-data' }],
            page: {
              paper: { type: 'A4' },
              orientation: 'landscape',
              margins: { top: 24, right: 20, bottom: 24, left: 20 },
              scale: { type: 'fixed', value: 0.9 },
            },
            manualBreaks: [{ sheetId: 'sheet-data', beforeRow: 2 }],
            header: { center: 'Semantic round trip' },
            footer: { right: 'Page &P' },
            showGridlines: false,
            showHeadings: true,
          },
        ],
      },
    ],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error('semantic fixture must be valid');
  return parsed.document;
}

describe('IO-01 semantic workbook writers', () => {
  it('exposes a backward-compatible blob write and a structured XLSX export result', async () => {
    const writer: WorkbookWriter = createXlsxWriter();
    const blob = await writer.write(semanticDocument());
    const result: WorkbookExportResult = await writer.writeResult(semanticDocument());

    expect(writer.format).toBe('xlsx');
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(result).toMatchObject({ format: 'xlsx', blob: expect.any(Blob), diagnostics: [] });
    expect(Object.isFrozen(result)).toBe(true);
  });

  it('round-trips supported XLSX values, formulas, styles, merges, validation, formatting, filter, and print', async () => {
    const source = semanticDocument();
    const exported = await createXlsxWriter().writeResult(source);
    expect(Object.keys(unzipSync(new Uint8Array(await exported.blob.arrayBuffer())))).not.toEqual(
      expect.arrayContaining([expect.stringMatching(/tego-sheet|customXml/i)]),
    );
    const imported = await createXlsxReader().read(exported.blob);
    const sheet = imported.document.workbook.sheets[0]!;

    expect(sheet.cells.map(({ cell }) => cell.input)).toEqual(
      source.workbook.sheets[0]!.cells.map(({ cell }) => cell.input),
    );
    expect(sheet.merges).toEqual(source.workbook.sheets[0]!.merges);
    expect(sheet.filter).toEqual(source.workbook.sheets[0]!.filter);
    expect(sheet.conditionalFormatting[0]).toMatchObject({
      ...source.workbook.sheets[0]!.conditionalFormatting[0],
      range: {
        ...source.workbook.sheets[0]!.conditionalFormatting[0]!.range,
        sheetId: sheet.id,
      },
    });
    const importedStyle = imported.document.workbook.styles.find(
      ({ id }) => id === sheet.cells[0]!.cell.styleId,
    );
    expect(importedStyle?.value).toMatchObject(
      source.workbook.styles[0]!.value as Record<string, unknown>,
    );
    const importedValidation = imported.document.workbook.validations.find(
      ({ id }) => id === sheet.cells[0]!.cell.validationId,
    );
    expect(importedValidation?.value).toEqual(source.workbook.validations[0]!.value);
    expect(imported.document.templates[0]?.printProfiles[0]).toMatchObject({
      page: { orientation: 'landscape' },
      header: { center: 'Semantic round trip' },
      showGridlines: false,
    });
    expect(imported.security.unsupportedFeatures).not.toEqual(
      expect.arrayContaining([
        'xlsx:styles',
        'xlsx:merged-cells',
        'xlsx:data-validation',
        'xlsx:conditional-formatting',
        'xlsx:auto-filter',
        'xlsx:print-settings',
      ]),
    );
  });

  it('writes deterministic bounded ODS and round-trips its supported semantic foundation', async () => {
    const source = semanticDocument();
    const writer = createOdsWriter();
    const first = await writer.writeResult(source);
    const second = await writer.writeResult(source);

    expect(new Uint8Array(await first.blob.arrayBuffer())).toEqual(
      new Uint8Array(await second.blob.arrayBuffer()),
    );
    const archive = unzipSync(new Uint8Array(await first.blob.arrayBuffer()));
    expect(new TextDecoder().decode(archive['content.xml'])).not.toMatch(
      /xmlns:tego|<tego:|\stego:/i,
    );
    const imported = await createOdsReader().read(first.blob);
    const sheet = imported.document.workbook.sheets[0]!;
    expect(sheet.cells.map(({ cell }) => cell.input)).toEqual(
      source.workbook.sheets[0]!.cells.map(({ cell }) => cell.input),
    );
    expect(sheet.merges).toEqual(source.workbook.sheets[0]!.merges);
    const importedStyle = imported.document.workbook.styles.find(
      ({ id }) => id === sheet.cells[0]!.cell.styleId,
    );
    expect(importedStyle?.value).toMatchObject(
      source.workbook.styles[0]!.value as Record<string, unknown>,
    );
    const importedValidation = imported.document.workbook.validations.find(
      ({ id }) => id === sheet.cells[0]!.cell.validationId,
    );
    expect(importedValidation?.value).toEqual(source.workbook.validations[0]!.value);
    expect(imported.document.templates[0]?.printProfiles[0]).toMatchObject({
      page: { orientation: 'landscape' },
      showGridlines: false,
    });
    expect(first.diagnostics.map(({ details }) => details)).toEqual(
      expect.arrayContaining([
        { feature: 'ods:auto-filter' },
        { feature: 'ods:conditional-formatting' },
      ]),
    );
  });

  it('honors abort and output quotas atomically for package writers', async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      createXlsxWriter().write(semanticDocument(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: 'ABORTED' });
    await expect(
      createOdsWriter({ maxOutputBytes: 16 }).write(semanticDocument()),
    ).rejects.toMatchObject({ code: 'OUTPUT_LIMIT_EXCEEDED' });
  });
});
