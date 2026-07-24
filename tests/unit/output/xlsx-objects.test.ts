import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { XlsxAdapter } from '../../../src/output/xlsx';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

describe('XlsxAdapter DrawingML objects', () => {
  it('writes standard one-cell shapes with EMU geometry and bounded object counts', async () => {
    const fixture = outputGeneratedDocument();
    const sheet = fixture.workbook.sheets[0]!;
    const document = {
      ...fixture,
      workbook: {
        ...fixture.workbook,
        sheets: [
          {
            ...sheet,
            objects: [
              {
                id: 'rectangle',
                kind: 'shape',
                shape: 'rectangle',
                anchor: {
                  type: 'one-cell',
                  cell: { sheetId: sheet.id, row: 1, column: 2 },
                  offset: { x: 3, y: 4 },
                  size: { width: 30, height: 40 },
                },
                zIndex: 0,
                locked: true,
                templateRepeat: 'shared',
                rotation: 15,
                style: { fill: '#112233', stroke: '#445566', strokeWidth: 2 },
                accessibility: { name: 'Accessible rectangle' },
              },
            ],
          },
        ],
      },
      objects: [],
    };
    const blob = await new XlsxAdapter().render(document as never, {
      formulaMode: 'values-only',
      compatibility: 'excel',
    });
    const entries = unzipSync(new Uint8Array(await blob.arrayBuffer()));
    const drawing = strFromU8(entries['xl/drawings/drawing1.xml']!);

    expect(drawing).toContain('<xdr:oneCellAnchor>');
    expect(drawing).toContain('<xdr:colOff>28575</xdr:colOff>');
    expect(drawing).toContain('<xdr:rowOff>38100</xdr:rowOff>');
    expect(drawing).toContain('<xdr:ext cx="285750" cy="381000"/>');
    expect(drawing).toContain('rot="900000"');
    expect(drawing).toContain('prst="rect"');
    expect(drawing).toContain('name="Accessible rectangle"');
    expect(drawing).toContain('<a:spLocks noMove="1" noResize="1"/>');

    await expect(
      new XlsxAdapter({ limits: { maxImages: 0 } }).render(document as never, {
        formulaMode: 'values-only',
        compatibility: 'excel',
      }),
    ).rejects.toMatchObject({ code: 'XLSX_PACKAGE_LIMIT_EXCEEDED' });
  });
});
