import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import {
  createCsvReader,
  createCsvWriter,
  createOdsReader,
  createTsvReader,
  createTsvWriter,
  createXlsxReader,
  InterchangeError,
  type WorkbookImportResult,
  type WorkbookReader,
  type WorkbookWriter,
} from '../../../src/interchange';
import { createSpreadsheetDocument } from '../../../src/document';

function archive(parts: Readonly<Record<string, string>>): Uint8Array {
  return zipSync(
    Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, strToU8(value)])),
  );
}

function xlsxFixture(
  worksheet = `<?xml version="1.0"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData><row r="1">
        <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
        <c r="B1"><v>42</v></c>
        <c r="C1" t="b"><v>1</v></c>
        <c r="D1"><f>B1+1</f><v>43</v></c>
      </row></sheetData>
    </worksheet>`,
  extra: Readonly<Record<string, string>> = {},
): Uint8Array {
  return archive({
    '[Content_Types].xml': '<Types/>',
    'xl/workbook.xml': `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': worksheet,
    ...extra,
  });
}

function odsFixture(
  content = `<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:version="1.3">
    <office:body><office:spreadsheet>
      <table:table table:name="Data">
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>Hello</text:p></table:table-cell>
          <table:table-cell office:value-type="float" office:value="42"/>
          <table:table-cell office:value-type="boolean" office:boolean-value="true"/>
        </table:table-row>
      </table:table>
    </office:spreadsheet></office:body>
  </office:document-content>`,
): Uint8Array {
  return archive({
    mimetype: 'application/vnd.oasis.opendocument.spreadsheet',
    'content.xml': content,
  });
}

describe('IO-01 bounded atomic workbook readers and writers', () => {
  it('round-trips CSV quoting, newlines, typed values, and formula-injection protection', async () => {
    const reader = createCsvReader();
    const imported = await reader.read(
      new TextEncoder().encode('name,amount,active\r\n"A, Inc",12,TRUE\r\n"two\nlines",,FALSE\r\n'),
    );
    expect(imported.document.workbook.sheets[0]?.cells.map(({ cell }) => cell.input)).toEqual([
      { type: 'string', value: 'name' },
      { type: 'string', value: 'amount' },
      { type: 'string', value: 'active' },
      { type: 'string', value: 'A, Inc' },
      { type: 'number', value: 12 },
      { type: 'boolean', value: true },
      { type: 'string', value: 'two\nlines' },
      { type: 'boolean', value: false },
    ]);
    expect(Object.isFrozen(imported)).toBe(true);
    expect(Object.isFrozen(imported.security)).toBe(true);
    expect(imported.security).toMatchObject({
      activeContentExecuted: false,
      externalResourcesFetched: false,
    });
    const writer = createCsvWriter();
    const protectedBlob = await writer.write(
      {
        ...imported.document,
        workbook: {
          ...imported.document.workbook,
          sheets: [
            {
              ...imported.document.workbook.sheets[0]!,
              cells: [
                {
                  row: 0,
                  column: 0,
                  cell: { input: { type: 'string', value: '=SUM(A1:A2)' } },
                },
              ],
            },
          ],
        },
      },
      { delimiter: ',', lineEnding: '\r\n', formulaInjectionProtection: true },
    );
    await expect(protectedBlob.text()).resolves.toBe(`'=SUM(A1:A2)\r\n`);
  });

  it('supports TSV with the same reader and writer contracts', async () => {
    const reader: WorkbookReader = createTsvReader();
    const imported: WorkbookImportResult = await reader.read(
      new TextEncoder().encode('name\tvalue\n"tab\tinside"\t-2.5\n'),
    );
    expect(imported.format).toBe('tsv');
    expect(imported.document.workbook.sheets[0]?.cells[2]?.cell.input).toEqual({
      type: 'string',
      value: 'tab\tinside',
    });
    const writer: WorkbookWriter = createTsvWriter();
    const output = await writer.write(imported.document);
    await expect(output.text()).resolves.toContain('"tab\tinside"\t-2.5');
  });

  it('imports a restricted XLSX core without executing active content or fetching relationships', async () => {
    const imported = await createXlsxReader().read(xlsxFixture());
    expect(imported.document.workbook.sheets[0]).toMatchObject({
      name: 'Data',
      cells: [
        { row: 0, column: 0, cell: { input: { type: 'string', value: 'Hello' } } },
        { row: 0, column: 1, cell: { input: { type: 'number', value: 42 } } },
        { row: 0, column: 2, cell: { input: { type: 'boolean', value: true } } },
        { row: 0, column: 3, cell: { input: { type: 'formula', source: '=B1+1' } } },
      ],
    });
    expect(imported.security).toEqual({
      activeContentExecuted: false,
      externalResourcesFetched: false,
      warnings: [],
      unsupportedFeatures: [],
    });
  });

  it('imports a restricted ODS core including merged ranges', async () => {
    const imported = await createOdsReader().read(
      odsFixture(
        odsFixtureContent(
          '<table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>Merged</text:p></table:table-cell>',
        ),
      ),
    );
    expect(imported.document.workbook.sheets[0]?.cells.map(({ cell }) => cell.input)).toEqual([
      { type: 'string', value: 'Merged' },
    ]);
    expect(imported.document.workbook.sheets[0]?.merges).toEqual([
      { start: { row: 0, column: 0 }, end: { row: 0, column: 1 } },
    ]);
    expect(imported.security.unsupportedFeatures).not.toContain('ods:merged-cells');
  });

  it('reports every recognized XLSX degradation as a structured diagnostic', async () => {
    const imported = await createXlsxReader().read(
      xlsxFixture(`<?xml version="1.0"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
          <autoFilter ref="A1:A2"/>
          <hyperlinks><hyperlink ref="A1" location="Data!A1"/></hyperlinks>
          <sheetProtection sheet="1"/>
          <drawing r:id="rId2"/>
          <tableParts count="1"><tablePart r:id="rId3"/></tableParts>
        </worksheet>`),
    );

    expect(imported.security.unsupportedFeatures).toEqual([
      'xlsx:hyperlinks',
      'xlsx:sheet-protection',
      'xlsx:drawing-objects',
      'xlsx:tables',
    ]);
    expect(imported.document.workbook.sheets[0]?.filter).toEqual({
      range: { start: { row: 0, column: 0 }, end: { row: 1, column: 0 } },
      filters: [],
    });
    expect(imported.diagnostics).toEqual(
      imported.security.unsupportedFeatures.map((feature) =>
        expect.objectContaining({
          code: 'UNSUPPORTED_INTERCHANGE_FEATURE',
          severity: 'warning',
          domain: 'interchange',
          stage: 'decode',
          details: { feature },
        }),
      ),
    );
  });

  it('reports native XLSX sparkline extensions as unsupported', async () => {
    const imported = await createXlsxReader().read(
      xlsxFixture(`<?xml version="1.0"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
          <extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups><x14:sparklineGroup>
              <x14:sparklines><x14:sparkline><xm:f>A1:A1</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline></x14:sparklines>
            </x14:sparklineGroup></x14:sparklineGroups>
          </ext></extLst>
        </worksheet>`),
    );

    expect(imported.security.unsupportedFeatures).toContain('xlsx:sparklines');
    expect(imported.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'UNSUPPORTED_INTERCHANGE_FEATURE',
        details: { feature: 'xlsx:sparklines' },
      }),
    );
  });

  it('rejects worksheet filters whose filter or sort column is outside the filter range', async () => {
    const fixture = (body: string) =>
      xlsxFixture(`<?xml version="1.0"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
          <autoFilter ref="A1:B2">${body}</autoFilter>
        </worksheet>`);

    await expect(
      createXlsxReader().read(
        fixture('<filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn>'),
      ),
    ).rejects.toMatchObject({ code: 'MALFORMED_WORKBOOK' });
    await expect(
      createXlsxReader().read(
        fixture('<sortState ref="A1:B2"><sortCondition ref="C1:C2"/></sortState>'),
      ),
    ).rejects.toMatchObject({ code: 'MALFORMED_WORKBOOK' });
  });

  it('reports every recognized ODS degradation as a structured diagnostic', async () => {
    const imported = await createOdsReader().read(
      odsFixture(`<?xml version="1.0"?>
        <office:document-content
          xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
          xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
          office:version="1.3">
          <office:body><office:spreadsheet>
            <table:named-expressions/>
            <table:content-validations/>
            <table:database-ranges/>
            <table:data-pilot-tables/>
            <table:table table:name="Data">
              <table:shapes><draw:frame/></table:shapes>
              <table:table-row>
                <table:table-cell office:value-type="string"><text:p>Hello</text:p></table:table-cell>
              </table:table-row>
            </table:table>
          </office:spreadsheet></office:body>
        </office:document-content>`),
    );

    expect(imported.security.unsupportedFeatures).toEqual([
      'ods:drawing-objects',
      'ods:database-ranges',
      'ods:named-expressions',
      'ods:pivot-tables',
    ]);
    expect(imported.diagnostics).toHaveLength(imported.security.unsupportedFeatures.length);
    expect(imported.diagnostics.every((entry) => entry.severity === 'warning')).toBe(true);
  });

  it('rejects archive and cell limits without exposing partial documents', async () => {
    const oversized = new Uint8Array(32);
    for (const reader of [
      createXlsxReader({ maxPackageBytes: 16 }),
      createOdsReader({ maxPackageBytes: 16 }),
    ]) {
      let failure: unknown;
      try {
        await reader.read(oversized);
      } catch (error) {
        failure = error;
      }
      expect(failure).toBeInstanceOf(InterchangeError);
      expect(failure).toMatchObject({ code: 'ARCHIVE_LIMIT_EXCEEDED' });
      expect('partialDocument' in Object(failure)).toBe(false);
    }
    await expect(
      createCsvReader({ maxCells: 1 }).read(new TextEncoder().encode('one,two')),
    ).rejects.toMatchObject({ code: 'CELL_LIMIT_EXCEEDED' });
  });

  it('rejects ZIP bombs, macros, external relationships, and XML entities atomically', async () => {
    await expect(
      createXlsxReader({ maxUncompressedBytes: 64 }).read(xlsxFixture()),
    ).rejects.toMatchObject({ code: 'ARCHIVE_LIMIT_EXCEEDED' });
    await expect(
      createXlsxReader().read(xlsxFixture(undefined, { 'xl/vbaProject.bin': 'macro' })),
    ).rejects.toMatchObject({ code: 'ACTIVE_CONTENT_REJECTED' });
    await expect(
      createXlsxReader().read(
        xlsxFixture(undefined, {
          'xl/_rels/workbook.xml.rels':
            '<Relationships><Relationship Id="rId1" TargetMode="External" Target="https://example.com/book.xlsx"/></Relationships>',
        }),
      ),
    ).rejects.toMatchObject({ code: 'EXTERNAL_RESOURCE_REJECTED' });
    await expect(
      createOdsReader().read(
        odsFixture('<!DOCTYPE x [<!ENTITY leak SYSTEM "file:///etc/passwd">]><x>&leak;</x>'),
      ),
    ).rejects.toMatchObject({ code: 'XML_ENTITY_REJECTED' });
  });

  it('honors AbortSignal before and during bounded reads without partial results', async () => {
    const alreadyAborted = new AbortController();
    alreadyAborted.abort();
    for (const reader of [createCsvReader(), createXlsxReader(), createOdsReader()]) {
      await expect(
        reader.read(new Uint8Array(), { signal: alreadyAborted.signal }),
      ).rejects.toMatchObject({ code: 'ABORTED' });
    }

    const controller = new AbortController();
    const input = new Blob(['a,b\n1,2\n']);
    const read = createCsvReader().read(input, { signal: controller.signal });
    controller.abort();
    await expect(read).rejects.toMatchObject({ code: 'ABORTED' });
  });

  it('rejects malformed quoted fields and text limits atomically', async () => {
    await expect(
      createCsvReader().read(new TextEncoder().encode('"unterminated')),
    ).rejects.toMatchObject({ code: 'MALFORMED_DELIMITED_TEXT' });
    await expect(
      createCsvReader().read(new TextEncoder().encode('unquoted"quote')),
    ).rejects.toMatchObject({ code: 'MALFORMED_DELIMITED_TEXT' });
    await expect(createCsvReader().read(new Uint8Array([0xc3, 0x28]))).rejects.toMatchObject({
      code: 'MALFORMED_DELIMITED_TEXT',
      security: {
        activeContentExecuted: false,
        externalResourcesFetched: false,
      },
    });
    await expect(
      createCsvReader({ maxFieldBytes: 3 }).read(new TextEncoder().encode('four')),
    ).rejects.toMatchObject({ code: 'FIELD_LIMIT_EXCEEDED' });
  });

  it('rejects macro-enabled XLSX content types even when the binary part is absent', async () => {
    await expect(
      createXlsxReader().read(
        xlsxFixture(undefined, {
          '[Content_Types].xml':
            '<Types><Override ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>',
        }),
      ),
    ).rejects.toMatchObject({
      code: 'ACTIVE_CONTENT_REJECTED',
      security: {
        activeContentExecuted: false,
        externalResourcesFetched: false,
      },
    });
  });

  it('writers consume immutable SpreadsheetDocument snapshots', async () => {
    const document = createSpreadsheetDocument({ id: 'document-1', sheetId: 'sheet-1' });
    const blob = await createCsvWriter().write(document, {
      delimiter: '\t',
      lineEnding: '\n',
      formulaInjectionProtection: true,
    });
    expect(blob.type).toBe('text/tab-separated-values');
    expect(document.workbook.sheets[0]?.cells).toEqual([]);
  });
});

function odsFixtureContent(cell: string): string {
  return `<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:version="1.3">
    <office:body><office:spreadsheet><table:table table:name="Data">
      <table:table-row>${cell}</table:table-row>
    </table:table></office:spreadsheet></office:body>
  </office:document-content>`;
}
