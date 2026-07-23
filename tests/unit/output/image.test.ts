import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { ImageAdapter, type ImageRasterizer } from '../../../src/output/image';
import { outputGeneratedDocument } from '../../fixtures/output/generated-document';

function withFontFsType(bytes: readonly number[], fsType: number): number[] {
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
  throw new Error('Image font fixture must contain an OS/2 table');
}

function imageDocument() {
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
  return {
    ...fixture,
    print: {
      ...fixture.print,
      displayList: {
        ...fixture.print.displayList,
        pages: fixture.print.displayList.pages.map((page, pageIndex) =>
          pageIndex === 0
            ? {
                ...page,
                commands: page.commands.map((command) =>
                  command.kind === 'text' ? { ...command, text: '中文' } : command,
                ),
              }
            : page,
        ),
      },
    },
    resources: {
      ...fixture.resources,
      byHash: { [resource.contentHash]: resource },
      byReference: { noto: resource },
      totalBytes: bytes.length,
    },
  };
}

describe('ImageAdapter', () => {
  it('emits requested standalone SVG pages in exact order and without active content', async () => {
    const blobs = await new ImageAdapter().render(imageDocument(), {
      format: 'svg',
      pages: [1, 0],
      background: 'transparent',
    });
    const markup = await Promise.all(blobs.map((blob) => blob.text()));

    expect(blobs.map((blob) => blob.type)).toEqual(['image/svg+xml', 'image/svg+xml']);
    expect(markup[0]).toContain('data-page-id="invoice-2"');
    expect(markup[1]).toContain('viewBox="0 0 210 297"');
    expect(markup[1]).toContain('<path');
    expect(markup.join('')).not.toContain('<style');
    expect(markup.join('')).not.toContain('base64');
    expect(markup.join('')).not.toMatch(/<script|foreignObject|\s(?:src|href)=["']https?:/iu);
  });

  it.each([
    [96, 210, 297],
    [150, 328, 464],
    [300, 656, 928],
  ])('calculates PNG pixels at %i DPI before rasterization', async (dpi, width, height) => {
    const rasterize = vi.fn(async () => new Blob([new Uint8Array([137, 80, 78, 71])]));
    const blobs = await new ImageAdapter({ rasterize }).render(imageDocument(), {
      format: 'png',
      pages: [0],
      dpi,
      background: '#ffffff',
    });

    expect(blobs[0]?.type).toBe('image/png');
    expect(rasterize).toHaveBeenCalledWith(
      expect.objectContaining({ width, height, dpi, background: '#ffffff' }),
    );
  });

  it('uses glyph outlines without exposing font-family CSS and rejects restricted fonts', async () => {
    const fixture = imageDocument();
    const resource = Object.values(fixture.resources.byHash)[0]!;
    const injectedFamily = 'Noto Sans";}@import url(https://example.com/font);/*';
    const injected = {
      ...fixture,
      print: {
        ...fixture.print,
        displayList: {
          ...fixture.print.displayList,
          pages: fixture.print.displayList.pages.map((page, pageIndex) =>
            pageIndex === 0
              ? {
                  ...page,
                  commands: page.commands.map((command) =>
                    command.kind === 'text' ? { ...command, fontFamily: injectedFamily } : command,
                  ),
                }
              : page,
          ),
        },
      },
      resources: {
        ...fixture.resources,
        byHash: {
          [resource.contentHash]: { ...resource, fontFamily: injectedFamily },
        },
      },
    };
    const [blob] = await new ImageAdapter().render(injected, {
      format: 'svg',
      pages: [0],
    });
    const svg = await blob!.text();

    expect(svg).toContain('<path');
    expect(svg).not.toContain(injectedFamily);
    expect(svg).not.toMatch(/<style|@import|\s(?:src|href|xlink:href)=["']https?:/iu);

    const restricted = {
      ...fixture,
      resources: {
        ...fixture.resources,
        byHash: {
          [resource.contentHash]: {
            ...resource,
            bytes: withFontFsType(resource.bytes, 0x0002),
          },
        },
      },
    };
    await expect(
      new ImageAdapter().render(restricted, { format: 'svg', pages: [0] }),
    ).rejects.toMatchObject({ code: 'IMAGE_FONT_EMBEDDING_FAILED' });
  });

  it('rejects the total pixel budget before allocating a raster surface', async () => {
    const rasterize = vi.fn(async () => new Blob());

    await expect(
      new ImageAdapter({ rasterize, limits: { maxPixels: 1_000 } }).render(imageDocument(), {
        format: 'png',
        pages: [0],
        dpi: 300,
      }),
    ).rejects.toMatchObject({ code: 'IMAGE_PIXEL_LIMIT_EXCEEDED' });
    expect(rasterize).not.toHaveBeenCalled();
  });

  it('reports a stable capability error when no PNG raster surface exists', async () => {
    await expect(
      new ImageAdapter().render(imageDocument(), {
        format: 'png',
        pages: [0],
        dpi: 96,
      }),
    ).rejects.toMatchObject({
      code: 'IMAGE_ENCODING_FAILED',
      message: 'PNG encoding failed',
    });
  });

  it('rejects invalid selection, aborts atomically, and returns no partial array', async () => {
    await expect(
      new ImageAdapter().render(imageDocument(), { format: 'svg', pages: [3] }),
    ).rejects.toMatchObject({ code: 'IMAGE_PAGE_SELECTION_INVALID' });

    const controller = new AbortController();
    let calls = 0;
    const rasterize = vi.fn(async () => {
      calls += 1;
      controller.abort();
      return new Blob([String(calls)]);
    });
    await expect(
      new ImageAdapter({ rasterize }).render(imageDocument(), {
        format: 'png',
        pages: [0, 1],
        dpi: 96,
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    expect(rasterize).toHaveBeenCalledTimes(1);

    const failOnSecondPage = vi
      .fn<ImageRasterizer>()
      .mockResolvedValueOnce(new Blob(['complete-first-page']))
      .mockRejectedValueOnce(new Error('second page failed'));
    await expect(
      new ImageAdapter({ rasterize: failOnSecondPage }).render(imageDocument(), {
        format: 'png',
        pages: [0, 1],
        dpi: 96,
      }),
    ).rejects.toMatchObject({ code: 'IMAGE_ENCODING_FAILED' });
    expect(failOnSecondPage).toHaveBeenCalledTimes(2);
  });
});
