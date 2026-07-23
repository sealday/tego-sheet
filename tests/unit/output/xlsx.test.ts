import { describe, expect, it } from 'vitest';
import { XlsxAdapter } from '../../../src/output/xlsx';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

async function parts(blob: Blob): Promise<Readonly<Record<string, string>>> {
  const { strFromU8, unzipSync } = await import('fflate');
  const archive = unzipSync(new Uint8Array(await blob.arrayBuffer()));
  return Object.fromEntries(
    Object.entries(archive).map(([path, bytes]) => [path, strFromU8(bytes)]),
  );
}

async function centralDirectory(
  blob: Blob,
): Promise<readonly { readonly name: string; readonly time: number; readonly date: number }[]> {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const entries: { name: string; time: number; date: number }[] = [];
  for (let offset = 0; offset <= bytes.length - 46; offset += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) continue;
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    entries.push({
      name: new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + nameLength)),
      time: view.getUint16(offset + 12, true),
      date: view.getUint16(offset + 14, true),
    });
    offset += 45 + nameLength + extraLength + commentLength;
  }
  return entries;
}

describe('XlsxAdapter', () => {
  it('writes typed cells, literal equals text, formulas with caches, styles, merges, validation, dimensions, and print parts', async () => {
    const source = outputGeneratedDocument();
    const blob = await new XlsxAdapter().render(source, {
      formulaMode: 'formula-and-cached-value',
      compatibility: 'excel',
    });
    const xml = await parts(blob);
    const sheet = xml['xl/worksheets/sheet1.xml'] ?? '';

    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(sheet).toContain('<c r="A1"');
    expect(sheet).toContain('>=literal</t>');
    expect(sheet).toContain('<f>A2+1</f><v>3</v>');
    expect(sheet).toContain('<mergeCell ref="A1:B1"/>');
    expect(sheet).toContain('<dataValidation');
    expect(sheet).toContain('<col min="2" max="2" width="80"');
    expect(sheet).toContain('hidden="1"');
    expect(sheet).toContain('<row r="1" ht="24" customHeight="1"');
    expect(sheet).toContain('<pageMargins');
    expect(sheet).toContain('<pageSetup');
    expect(sheet).toContain('<rowBreaks');
    expect(sheet).toContain('<headerFooter>');
    expect(xml['xl/styles.xml']).toContain('<numFmt');
    expect(xml['xl/workbook.xml']).toContain('_xlnm.Print_Titles');
  });

  it('removes formula markup in values-only mode', async () => {
    const xml = await parts(
      await new XlsxAdapter().render(outputGeneratedDocument(), {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    );

    expect(xml['xl/worksheets/sheet1.xml']).not.toContain('<f>');
  });

  it('marks cached formula errors with the OOXML error type', async () => {
    const fixture = outputGeneratedDocument();
    const xml = await parts(
      await new XlsxAdapter().render(
        {
          ...fixture,
          calculatedCells: [
            {
              address: { sheetId: 'sheet-1', row: 1, column: 1 },
              value: { type: 'error', value: '#DIV/0!' },
            },
          ],
        } as never,
        { formulaMode: 'formula-and-cached-value', compatibility: 'excel' },
      ),
    );

    expect(xml['xl/worksheets/sheet1.xml']).toContain(
      '<c r="B2" t="e"><f>A2+1</f><v>#DIV/0!</v></c>',
    );
  });

  it('preserves exact range print areas and custom paper geometry', async () => {
    const fixture = outputGeneratedDocument();
    const blob = await new XlsxAdapter().render(
      {
        ...fixture,
        print: {
          ...fixture.print,
          profile: {
            ...fixture.print.profile,
            targets: [
              {
                type: 'range',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 1, column: 1 },
                  end: { row: 2, column: 1 },
                },
              },
            ],
            page: {
              ...fixture.print.profile.page,
              paper: { type: 'custom', width: 384, height: 192 },
            },
          },
        },
      } as never,
      { formulaMode: 'values-only', compatibility: 'excel' },
    );
    const xml = await parts(blob);

    expect(xml['xl/workbook.xml']).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0">'Invoice'!$B$2:$B$3</definedName>`,
    );
    expect(xml['xl/worksheets/sheet1.xml']).toContain('paperWidth="4in" paperHeight="2in"');
  });

  it('produces byte-identical ZIP output for equal inputs and options', async () => {
    const adapter = new XlsxAdapter();
    const options = {
      formulaMode: 'values-only' as const,
      compatibility: 'excel' as const,
    };

    const first = new Uint8Array(
      await (await adapter.render(outputGeneratedDocument(), options)).arrayBuffer(),
    );
    const second = new Uint8Array(
      await (await adapter.render(outputGeneratedDocument(), options)).arrayBuffer(),
    );
    expect(second).toEqual(first);
    const directory = await centralDirectory(new Blob([first]));
    expect(directory.map(({ name }) => name)).toEqual(directory.map(({ name }) => name).sort());
    expect(new Set(directory.map(({ time, date }) => `${time}:${date}`))).toEqual(
      new Set(['0:33']),
    );
  });

  it('serializes typed conditional formatting and worksheet visibility', async () => {
    const fixture = outputGeneratedDocument();
    const secondSheet = {
      ...fixture.workbook.sheets[0]!,
      id: 'sheet-2',
      name: 'Archive',
      cells: [],
      merges: [],
    };
    const xml = await parts(
      await new XlsxAdapter().render(
        {
          ...fixture,
          workbook: {
            ...fixture.workbook,
            sheets: [...fixture.workbook.sheets, secondSheet],
          },
          worksheets: [
            {
              sheetId: 'sheet-1',
              visibility: 'visible',
              conditionalFormatting: [
                {
                  type: 'color-scale',
                  range: {
                    sheetId: 'sheet-1',
                    start: { row: 1, column: 0 },
                    end: { row: 2, column: 1 },
                  },
                  minimumColor: '#ff0000',
                  midpointColor: '#ffff00',
                  maximumColor: '#00ff00',
                },
              ],
            },
            {
              sheetId: 'sheet-2',
              visibility: 'very-hidden',
              conditionalFormatting: [],
            },
          ],
        } as never,
        { formulaMode: 'values-only', compatibility: 'excel' },
      ),
    );

    expect(xml['xl/workbook.xml']).toContain(
      '<sheet name="Archive" sheetId="2" state="veryHidden" r:id="rId2"/>',
    );
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<conditionalFormatting sqref="A2:B3">');
    expect(xml['xl/worksheets/sheet1.xml']).toContain(
      '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>',
    );
    expect(xml['xl/worksheets/sheet1.xml']).toContain(
      '<color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/>',
    );
  });

  it('escapes XML and emits only internal image relationships', async () => {
    const fixture = outputGeneratedDocument();
    const image = {
      contentHash: 'sha256:logo',
      type: 'image' as const,
      mimeType: 'image/png',
      bytes: [137, 80, 78, 71],
    };
    const source = {
      ...fixture,
      workbook: {
        ...fixture.workbook,
        sheets: [
          {
            ...fixture.workbook.sheets[0]!,
            name: 'R&D <Q>',
            cells: [
              {
                row: 0,
                column: 0,
                cell: { input: { type: 'string', value: '<unsafe & =literal>' } },
              },
            ],
          },
        ],
      },
      objects: [
        {
          objectId: 'logo',
          resourceId: 'logo',
          policy: 'shared',
          itemIndex: 0,
          source: {
            sheetId: 'sheet-1',
            start: { row: 0, column: 0 },
            end: { row: 1, column: 1 },
          },
          generated: {
            sheetId: 'sheet-1',
            start: { row: 0, column: 0 },
            end: { row: 1, column: 1 },
          },
        },
      ],
      resources: {
        ...fixture.resources,
        byHash: { [image.contentHash]: image },
        byReference: { logo: image },
        totalBytes: image.bytes.length,
      },
    };
    const xmlParts = await parts(
      await new XlsxAdapter().render(source as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    );

    expect(xmlParts['xl/workbook.xml']).toContain('name="R&amp;D &lt;Q&gt;"');
    expect(xmlParts['xl/worksheets/sheet1.xml']).toContain('&lt;unsafe &amp; =literal&gt;');
    expect(xmlParts['xl/drawings/drawing1.xml']).toContain('r:embed="rId1"');
    expect(xmlParts['xl/drawings/_rels/drawing1.xml.rels']).toContain(
      'Target="../media/image1-1.png"',
    );
    expect(Object.values(xmlParts).join('')).not.toContain('TargetMode="External"');
  });

  it('rejects unsupported custom cells with a located diagnostic and no package', async () => {
    const source = outputGeneratedDocument();
    const custom = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: {
                    type: 'custom',
                    cellType: 'vendor/widget',
                    schemaVersion: 1,
                    value: {},
                  },
                },
              },
            ],
          },
        ],
      },
    };

    await expect(
      new XlsxAdapter().render(custom as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({
      code: 'XLSX_UNSUPPORTED_FEATURE',
      diagnostic: { location: { sheetId: 'sheet-1' } },
    });
  });

  it('diagnoses unsupported conditional formatting metadata instead of dropping it', async () => {
    const fixture = outputGeneratedDocument();
    const first = fixture.workbook.sheets[0]!.cells[0]!;
    const source = {
      ...fixture,
      workbook: {
        ...fixture.workbook,
        sheets: [
          {
            ...fixture.workbook.sheets[0]!,
            cells: [
              {
                ...first,
                cell: {
                  ...first.cell,
                  metadata: { conditionalFormatting: [{ type: 'colorScale' }] },
                },
              },
            ],
          },
        ],
      },
    };

    await expect(
      new XlsxAdapter().render(source as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({
      code: 'XLSX_UNSUPPORTED_FEATURE',
      diagnostic: { location: { sheetId: 'sheet-1' } },
    });
  });

  it('rejects non-allowlisted OOXML border and validation enums', async () => {
    const fixture = outputGeneratedDocument();
    const invalidBorder = {
      ...fixture,
      workbook: {
        ...fixture.workbook,
        styles: [
          {
            id: 'bad',
            value: { border: { left: ['javascript', '#000000'] } },
          },
        ],
      },
    };
    await expect(
      new XlsxAdapter().render(invalidBorder as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({ code: 'XLSX_UNSUPPORTED_FEATURE' });

    const invalidOperator = {
      ...fixture,
      workbook: {
        ...fixture.workbook,
        validations: [
          {
            id: 'positive',
            value: { type: 'decimal', operator: 'javascript', formula1: 0 },
          },
        ],
      },
    };
    await expect(
      new XlsxAdapter().render(invalidOperator as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({ code: 'XLSX_UNSUPPORTED_FEATURE' });
  });

  it('preflights source strings, resources, and uncompressed parts before publishing', async () => {
    const fixture = outputGeneratedDocument();
    const options = {
      formulaMode: 'values-only' as const,
      compatibility: 'excel' as const,
    };

    await expect(
      new XlsxAdapter({ limits: { maxStringBytes: 1 } }).render(fixture, options),
    ).rejects.toMatchObject({ code: 'XLSX_PACKAGE_LIMIT_EXCEEDED' });
    await expect(
      new XlsxAdapter({ limits: { maxResourceBytes: 1 } }).render(
        {
          ...fixture,
          resources: { ...fixture.resources, totalBytes: 2 },
        },
        options,
      ),
    ).rejects.toMatchObject({ code: 'XLSX_PACKAGE_LIMIT_EXCEEDED' });
    await expect(
      new XlsxAdapter({ limits: { maxUncompressedBytes: 1 } }).render(fixture, options),
    ).rejects.toMatchObject({ code: 'XLSX_PACKAGE_LIMIT_EXCEEDED' });
  });

  it('aborts before publishing a partial ZIP', async () => {
    const controller = new AbortController();
    controller.abort();

    await expect(
      new XlsxAdapter().render(outputGeneratedDocument(), {
        formulaMode: 'values-only',
        compatibility: 'excel',
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
  });
});
