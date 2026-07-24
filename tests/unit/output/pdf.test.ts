import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PdfAdapter } from '../../../src/output/pdf';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

async function pdfText(blob: Blob): Promise<string> {
  return new TextDecoder('latin1').decode(await blob.arrayBuffer());
}

function withFontFsType(bytes: Uint8Array, fsType: number): number[] {
  const copy = new Uint8Array(bytes);
  const view = new DataView(copy.buffer);
  const tableCount = view.getUint16(4);
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const tag = new TextDecoder().decode(copy.slice(record, record + 4));
    if (tag !== 'OS/2') continue;
    view.setUint16(view.getUint32(record + 8) + 8, fsType);
    return [...copy];
  }
  throw new Error('CJK font fixture must contain an OS/2 table');
}

describe('PdfAdapter', () => {
  it('translates grouped rotation and clipped shape commands into PDF graphics state', async () => {
    const fixture = outputGeneratedDocument();
    const page = fixture.print.displayList.pages[0]!;
    const blob = await new PdfAdapter().render(
      {
        ...fixture,
        print: {
          ...fixture.print,
          displayList: {
            diagnostics: [],
            pages: [
              {
                ...page,
                commands: [
                  {
                    kind: 'group',
                    rotation: 90,
                    origin: { x: 30, y: 35 },
                    commands: [
                      {
                        kind: 'clip',
                        rect: { x: 10, y: 20, width: 40, height: 30 },
                        commands: [
                          {
                            kind: 'fill-rect',
                            rect: { x: 10, y: 20, width: 40, height: 30 },
                            color: '#ffeecc',
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
              fixture.print.displayList.pages[1]!,
            ],
          },
        },
      } as never,
      { pages: [0], tagged: false },
    );
    const content = await pdfText(blob);

    expect(content).toContain('0 1 -1 0');
    expect(content.match(/\bq\b/g)?.length).toBeGreaterThanOrEqual(2);
    expect(content.match(/\bW\b/g)?.length).toBeGreaterThanOrEqual(1);
  });

  it('translates exact selected page geometry, vector content, searchable text, links, and metadata', async () => {
    const fixture = outputGeneratedDocument();
    const source = {
      ...fixture,
      print: {
        ...fixture.print,
        displayList: {
          ...fixture.print.displayList,
          pages: [
            {
              ...fixture.print.displayList.pages[0]!,
              commands: fixture.print.displayList.pages[0]!.commands.map((command) =>
                command.kind === 'text'
                  ? { ...command, text: 'Invoice', fontFamily: 'Helvetica' }
                  : command,
              ),
            },
            fixture.print.displayList.pages[1]!,
          ],
        },
      },
    };
    const blob = await new PdfAdapter().render(source, {
      pages: [0],
      tagged: false,
      metadata: { title: 'Invoice archive', author: 'Tego Sheet' },
    });
    const content = await pdfText(blob);

    expect(blob.type).toBe('application/pdf');
    expect(content.startsWith('%PDF-')).toBe(true);
    expect(content).toContain('/MediaBox [0 0 157.5 222.75]');
    expect(content).toContain('/Subtype /Link');
    expect(content).toContain('/Encoding /WinAnsiEncoding');
    expect(content).toContain('Invoice archive');
    expect(content.match(/\/Type \/Page\b/g)).toHaveLength(1);
  });

  it('rejects invalid page selections and unsupported commands without a partial Blob', async () => {
    const source = outputGeneratedDocument();
    await expect(
      new PdfAdapter().render(source, { pages: [2], tagged: false }),
    ).rejects.toMatchObject({ code: 'PDF_PAGE_SELECTION_INVALID' });

    const unsupported = {
      ...source,
      print: {
        ...source.print,
        displayList: {
          ...source.print.displayList,
          pages: [
            {
              ...source.print.displayList.pages[0]!,
              commands: [{ kind: 'execute-javascript', source: 'alert(1)' }],
            },
          ],
        },
      },
    };
    await expect(
      new PdfAdapter().render(unsupported as never, { pages: 'all', tagged: false }),
    ).rejects.toMatchObject({ code: 'PDF_UNSUPPORTED_DRAW_COMMAND' });
  });

  it('requires an embeddable matching font for CJK and never silently substitutes', async () => {
    await expect(
      new PdfAdapter().render(outputGeneratedDocument(), { pages: 'all', tagged: false }),
    ).rejects.toMatchObject({ code: 'PDF_FONT_SUBSET_FAILED' });
  });

  it('embeds a searchable CJK subset instead of outlining or replacing text', async () => {
    const fixture = outputGeneratedDocument();
    const encoded = readFileSync(
      new URL('../../fixtures/output/NotoSansSC-CJK.subset.ttf.base64', import.meta.url),
      'utf8',
    ).trim();
    const bytes = [...Buffer.from(encoded, 'base64')];
    const resource = {
      contentHash: 'sha256:noto-cjk-fixture',
      type: 'font' as const,
      mimeType: 'font/ttf',
      bytes,
      fontFamily: 'Noto Sans',
      fontEmbedding: 'allowed' as const,
    };
    const blob = await new PdfAdapter().render(
      {
        ...fixture,
        resources: {
          ...fixture.resources,
          byHash: { [resource.contentHash]: resource },
          byReference: { noto: resource },
          totalBytes: bytes.length,
        },
      },
      { pages: [0], tagged: false },
    );
    const content = await pdfText(blob);

    expect(content).toContain('/ToUnicode');
    expect(content).toContain('/FontFile2');
    expect(blob.size).toBeLessThan(100_000);
  });

  it('rejects restricted OS/2 font embedding before PDFKit receives the font', async () => {
    const fixture = outputGeneratedDocument();
    const encoded = readFileSync(
      new URL('../../fixtures/output/NotoSansSC-CJK.subset.ttf.base64', import.meta.url),
      'utf8',
    ).trim();
    const resource = {
      contentHash: 'sha256:restricted-cjk-fixture',
      type: 'font' as const,
      mimeType: 'font/ttf',
      bytes: withFontFsType(Buffer.from(encoded, 'base64'), 0x0002),
      fontFamily: 'Noto Sans',
    };

    await expect(
      new PdfAdapter().render(
        {
          ...fixture,
          resources: {
            ...fixture.resources,
            byHash: { [resource.contentHash]: resource },
            byReference: { noto: resource },
            totalBytes: resource.bytes.length,
          },
        },
        { pages: [0], tagged: false },
      ),
    ).rejects.toMatchObject({ code: 'PDF_FONT_EMBEDDING_FORBIDDEN' });
  });

  it('does not inspect or register fonts unused by the selected pages', async () => {
    const fixture = outputGeneratedDocument();
    const encoded = readFileSync(
      new URL('../../fixtures/output/NotoSansSC-CJK.subset.ttf.base64', import.meta.url),
      'utf8',
    ).trim();
    const resource = {
      contentHash: 'sha256:unused-restricted-font',
      type: 'font' as const,
      mimeType: 'font/ttf',
      bytes: withFontFsType(Buffer.from(encoded, 'base64'), 0x0002),
      fontFamily: 'Unused Restricted Font',
    };
    const page = fixture.print.displayList.pages[0]!;
    const blob = await new PdfAdapter().render(
      {
        ...fixture,
        resources: {
          ...fixture.resources,
          byHash: { [resource.contentHash]: resource },
          byReference: { unused: resource },
          totalBytes: resource.bytes.length,
        },
        print: {
          ...fixture.print,
          displayList: {
            ...fixture.print.displayList,
            pages: [
              {
                ...page,
                commands: page.commands.map((command) =>
                  command.kind === 'text'
                    ? { ...command, text: 'Invoice', fontFamily: 'Helvetica' }
                    : command,
                ),
              },
              fixture.print.displayList.pages[1]!,
            ],
          },
        },
      },
      { pages: [0], tagged: false },
    );

    expect(blob.type).toBe('application/pdf');
  });

  it('produces byte-identical output for equal inputs and metadata', async () => {
    const fixture = outputGeneratedDocument();
    const source = {
      ...fixture,
      print: {
        ...fixture.print,
        displayList: {
          diagnostics: [],
          pages: [
            {
              index: 0,
              width: 10,
              height: 10,
              commands: [],
            },
          ],
        },
        pages: [{ ...fixture.print.pages[0]!, width: 10, height: 10 }],
      },
    };
    const adapter = new PdfAdapter();
    const options = { pages: 'all' as const, tagged: false as const };

    const first = new Uint8Array(await (await adapter.render(source, options)).arrayBuffer());
    const second = new Uint8Array(await (await adapter.render(source, options)).arrayBuffer());

    expect(second).toEqual(first);
  });

  it('aborts atomically and allows a clean subsequent render attempt', async () => {
    const controller = new AbortController();
    controller.abort();
    const adapter = new PdfAdapter();

    await expect(
      adapter.render(outputGeneratedDocument(), {
        pages: 'all',
        tagged: false,
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    await expect(
      adapter.render(
        {
          ...outputGeneratedDocument(),
          print: {
            ...outputGeneratedDocument().print,
            displayList: {
              diagnostics: [],
              pages: [
                {
                  index: 0,
                  width: 10,
                  height: 10,
                  commands: [],
                },
              ],
            },
            pages: [
              {
                id: 'empty',
                index: 0,
                targetId: 'empty',
                width: 10,
                height: 10,
                rowStart: 0,
                rowEnd: 0,
                columnStart: 0,
                columnEnd: 0,
              },
            ],
          },
        },
        { pages: 'all', tagged: false },
      ),
    ).resolves.toBeInstanceOf(Blob);
  });

  it('enforces the wall-clock deadline for a single-page render', async () => {
    await expect(
      new PdfAdapter({ limits: { maxDurationMs: 0 } }).render(outputGeneratedDocument(), {
        pages: [0],
        tagged: false,
      }),
    ).rejects.toMatchObject({ code: 'PDF_OUTPUT_LIMIT_EXCEEDED' });
  });
});
