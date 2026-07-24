import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createXlsxReader, createXlsxWriter } from '../../../src/interchange';
import { XlsxAdapter } from '../../../src/output/xlsx';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

const imageBytes = Object.freeze([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]);

function objectDocument() {
  const fixture = outputGeneratedDocument();
  const sheet = fixture.workbook.sheets[0]!;
  const base = {
    zIndex: 0,
    locked: true,
    templateRepeat: 'shared' as const,
    accessibility: { name: 'Accessible object', description: 'Imported description' },
  };
  const objects = [
    {
      ...base,
      id: 'image-object',
      kind: 'image' as const,
      resourceId: 'logo',
      anchor: {
        type: 'absolute' as const,
        rect: { x: 1, y: 2, width: 10, height: 20 },
      },
      rotation: 15,
    },
    {
      ...base,
      id: 'rectangle-object',
      kind: 'shape' as const,
      shape: 'rectangle' as const,
      zIndex: 1,
      anchor: {
        type: 'one-cell' as const,
        cell: { sheetId: sheet.id, row: 1, column: 2 },
        offset: { x: 3, y: 4 },
        size: { width: 30, height: 40 },
      },
      style: { fill: '#112233', stroke: '#445566', strokeWidth: 2 },
    },
    {
      ...base,
      id: 'ellipse-object',
      kind: 'shape' as const,
      shape: 'ellipse' as const,
      zIndex: 2,
      locked: false,
      anchor: {
        type: 'two-cell' as const,
        from: { sheetId: sheet.id, row: 2, column: 1, offset: { x: 5, y: 6 } },
        to: { sheetId: sheet.id, row: 4, column: 3, offset: { x: 7, y: 8 } },
      },
      style: { fill: '#abcdef' },
    },
    {
      ...base,
      id: 'line-object',
      kind: 'shape' as const,
      shape: 'line' as const,
      zIndex: 3,
      anchor: {
        type: 'one-cell' as const,
        cell: { sheetId: sheet.id, row: 5, column: 0 },
        offset: { x: 0, y: 0 },
        size: { width: 50, height: 1 },
      },
      style: { stroke: '#ff0000', strokeWidth: 1 },
    },
    {
      ...base,
      id: 'text-object',
      kind: 'text-box' as const,
      zIndex: 4,
      anchor: {
        type: 'absolute' as const,
        rect: { x: 12, y: 13, width: 80, height: 24 },
      },
      text: 'Plain <text> & value',
      style: {
        color: '#123456',
        fontFamily: 'Arial',
        fontSize: 12,
        horizontalAlign: 'center' as const,
      },
    },
  ];
  const image = {
    contentHash: 'sha256:xlsx-object-image',
    type: 'image' as const,
    mimeType: 'image/png',
    bytes: imageBytes,
  };
  return {
    ...fixture,
    workbook: {
      ...fixture.workbook,
      sheets: [{ ...sheet, objects }],
    },
    resources: {
      ...fixture.resources,
      byHash: { [image.contentHash]: image },
      byReference: { logo: image },
      totalBytes: image.bytes.length,
    },
    objects: [],
  };
}

async function archiveParts(blob: Blob): Promise<Record<string, Uint8Array>> {
  return unzipSync(new Uint8Array(await blob.arrayBuffer()));
}

function repack(
  entries: Readonly<Record<string, Uint8Array>>,
  changes: Readonly<Record<string, string>>,
): Uint8Array {
  return zipSync({
    ...entries,
    ...Object.fromEntries(Object.entries(changes).map(([name, xml]) => [name, strToU8(xml)])),
  });
}

describe('OBJ-01 standard XLSX DrawingML interchange', () => {
  it('round-trips standard images, shapes, text boxes, anchors, and object metadata', async () => {
    const blob = await new XlsxAdapter().render(objectDocument() as never, {
      formulaMode: 'values-only',
      compatibility: 'excel',
    });
    const entries = await archiveParts(blob);
    const drawing = strFromU8(entries['xl/drawings/drawing1.xml']!);

    expect(drawing).toContain('<xdr:absoluteAnchor>');
    expect(drawing).toContain('<xdr:oneCellAnchor>');
    expect(drawing).toContain('<xdr:twoCellAnchor editAs="twoCell">');
    expect(drawing).toContain('<xdr:pos x="9525" y="19050"/>');
    expect(drawing).toContain('<xdr:ext cx="95250" cy="190500"/>');
    expect(drawing).toContain('rot="900000"');
    expect(drawing).toContain('prst="rect"');
    expect(drawing).toContain('prst="ellipse"');
    expect(drawing).toContain('prst="line"');
    expect(drawing).toContain('txBox="1"');
    expect(drawing).toContain('name="Accessible object"');
    expect(drawing).toContain('descr="Imported description"');
    expect(drawing.indexOf('<xdr:pic>')).toBeLessThan(drawing.indexOf('txBox="1"'));
    expect(strFromU8(entries['xl/drawings/_rels/drawing1.xml.rels']!)).toContain(
      'relationships/image',
    );
    expect(entries['xl/media/image1-1.png']).toEqual(new Uint8Array(imageBytes));
    expect(Object.keys(entries)).not.toEqual(
      expect.arrayContaining([expect.stringMatching(/customXml|tego-sheet/i)]),
    );

    const imported = await createXlsxReader().read(blob);
    const sheet = imported.document.workbook.sheets[0]!;
    expect(sheet.objects.map(({ kind }) => kind)).toEqual([
      'image',
      'shape',
      'shape',
      'shape',
      'text-box',
    ]);
    expect(sheet.objects.map(({ anchor }) => anchor.type)).toEqual([
      'absolute',
      'one-cell',
      'two-cell',
      'one-cell',
      'absolute',
    ]);
    expect(sheet.objects[0]).toMatchObject({
      kind: 'image',
      rotation: 15,
      locked: true,
      accessibility: {
        name: 'Accessible object',
        description: 'Imported description',
      },
    });
    expect(sheet.objects[1]).toMatchObject({
      kind: 'shape',
      shape: 'rectangle',
      style: { fill: '#112233', stroke: '#445566', strokeWidth: 2 },
    });
    expect(sheet.objects[4]).toMatchObject({
      kind: 'text-box',
      text: 'Plain <text> & value',
      style: { color: '#123456', fontFamily: 'Arial', fontSize: 12 },
    });
    expect(imported.document.resources.items[0]).toMatchObject({
      kind: 'image',
      mimeType: 'image/png',
      byteLength: imageBytes.length,
      url: expect.stringMatching(/^data:image\/png;base64,/),
    });
    expect(imported.security.unsupportedFeatures).not.toContain('xlsx:drawing-objects');

    const exportedAgain = await createXlsxWriter().writeResult(imported.document);
    const secondEntries = await archiveParts(exportedAgain.blob);
    expect(secondEntries['xl/drawings/drawing1.xml']).toBeDefined();
    expect(secondEntries['xl/media/image1-1.png']).toEqual(new Uint8Array(imageBytes));
    expect(exportedAgain.diagnostics).toEqual([]);
  });

  it.each([
    {
      name: 'external image relationship',
      feature: 'xlsx:drawing-external-relationship',
      mutate: (xml: string) =>
        xml.replace(
          'Target="../media/image1-1.png"',
          'Target="https://example.com/image.png" TargetMode="External"',
        ),
    },
    {
      name: 'package-escaping image relationship',
      feature: 'xlsx:drawing-resource-unsafe',
      mutate: (xml: string) =>
        xml.replace('Target="../media/image1-1.png"', 'Target="../../../image.png"'),
    },
  ])('degrades $name without fetching or exposing bytes', async ({ feature, mutate }) => {
    const blob = await new XlsxAdapter().render(objectDocument() as never, {
      formulaMode: 'values-only',
      compatibility: 'excel',
    });
    const entries = await archiveParts(blob);
    const relsPath = 'xl/drawings/_rels/drawing1.xml.rels';
    const hostile = repack(entries, { [relsPath]: mutate(strFromU8(entries[relsPath]!)) });

    const imported = await createXlsxReader().read(hostile);
    expect(imported.document.workbook.sheets[0]!.objects).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ kind: 'image' })]),
    );
    expect(imported.security.externalResourcesFetched).toBe(false);
    expect(imported.security.unsupportedFeatures).toContain(feature);
  });

  it('degrades unknown preset shapes and enforces DrawingML XML budgets', async () => {
    const blob = await new XlsxAdapter().render(objectDocument() as never, {
      formulaMode: 'values-only',
      compatibility: 'excel',
    });
    const entries = await archiveParts(blob);
    const drawingPath = 'xl/drawings/drawing1.xml';
    const drawing = strFromU8(entries[drawingPath]!);
    const unknown = repack(entries, {
      [drawingPath]: drawing.replaceAll('prst="rect"', 'prst="star5"'),
    });

    const imported = await createXlsxReader().read(unknown);
    expect(imported.security.unsupportedFeatures).toContain('xlsx:drawing-unknown-shape');
    expect(imported.document.workbook.sheets[0]!.objects).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ shape: 'rectangle' })]),
    );
    await expect(
      createXlsxReader({ maxXmlBytes: drawing.length - 1 }).read(blob),
    ).rejects.toMatchObject({ code: 'XML_LIMIT_EXCEEDED' });
    const totalUncompressedBytes = Object.values(entries).reduce(
      (total, bytes) => total + bytes.byteLength,
      0,
    );
    await expect(
      createXlsxReader({ maxUncompressedBytes: totalUncompressedBytes - 1 }).read(blob),
    ).rejects.toMatchObject({ code: 'ARCHIVE_LIMIT_EXCEEDED' });
    const entity = repack(entries, {
      [drawingPath]: `<!DOCTYPE x [<!ENTITY hostile "boom">]>${drawing}`,
    });
    await expect(createXlsxReader().read(entity)).rejects.toMatchObject({
      code: 'XML_ENTITY_REJECTED',
    });
  });

  it('counts standard worksheet objects against the bounded DrawingML budget', async () => {
    await expect(
      new XlsxAdapter({ limits: { maxImages: 4 } }).render(objectDocument() as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({ code: 'XLSX_PACKAGE_LIMIT_EXCEEDED' });
  });
});
