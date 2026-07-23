import { describe, expect, it } from 'vitest';
import { PdfAdapter } from '../../../src/output/pdf';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

async function pdfText(blob: Blob): Promise<string> {
  return new TextDecoder('latin1').decode(await blob.arrayBuffer());
}

describe('PdfAdapter', () => {
  it('translates exact selected page geometry, vector content, searchable text, links, and metadata', async () => {
    const source = outputGeneratedDocument();
    const blob = await new PdfAdapter().render(source, {
      pages: [0],
      tagged: false,
      metadata: { title: 'Invoice archive', author: 'Tego Sheet' },
    });
    const content = await pdfText(blob);

    expect(blob.type).toBe('application/pdf');
    expect(content.startsWith('%PDF-')).toBe(true);
    expect(content).toContain('/MediaBox [0 0 210 297]');
    expect(content).toContain('/Subtype /Link');
    expect(content).toContain('/ToUnicode');
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
});
