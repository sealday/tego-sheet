// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PrintDisplayList } from '../../../src/print';
import {
  BrowserPrintError,
  IsolatedBrowserPrintAdapter,
  serializeGeneratedDocumentSvgPages,
  type GeneratedDocumentForBrowserPrint,
} from '../../../src/output/browser-print-adapter';

function generatedDocument(): GeneratedDocumentForBrowserPrint {
  const pages: PrintDisplayList['pages'] = [
    {
      index: 0,
      width: 210,
      height: 297,
      commands: [
        {
          kind: 'fill-rect',
          rect: { x: 10, y: 20, width: 40, height: 12 },
          color: '#ffffff',
        },
        {
          kind: 'text',
          text: 'Invoice <not-markup>',
          x: 12,
          y: 27,
          maxWidth: 36,
          fontFamily: 'sans-serif',
          fontSize: 10,
          color: '#111111',
          horizontalAlign: 'left',
        },
      ],
    },
    {
      index: 1,
      width: 210,
      height: 297,
      commands: [
        {
          kind: 'line',
          x1: 10,
          y1: 10,
          x2: 200,
          y2: 10,
          color: '#222222',
          width: 0.5,
        },
      ],
    },
  ];
  const print: GeneratedDocumentForBrowserPrint['print'] = {
    pages: [
      { id: 'invoice-cover', index: 0, width: 210, height: 297 },
      { id: 'invoice-lines', index: 1, width: 210, height: 297 },
    ],
    displayList: { diagnostics: [], pages },
  };
  return { print };
}

function frame(): HTMLIFrameElement {
  const result = document.querySelector('iframe[data-tego-browser-print]');
  if (!(result instanceof HTMLIFrameElement)) throw new Error('print iframe was not mounted');
  return result;
}

afterEach(() => {
  vi.useRealTimers();
  document.body.replaceChildren();
  document.head.replaceChildren();
});

describe('IsolatedBrowserPrintAdapter', () => {
  it('serializes grouped rotation, vector shapes, and text clipping exactly', () => {
    const source = generatedDocument();
    const rotated = {
      ...source,
      print: {
        ...source.print,
        displayList: {
          diagnostics: [],
          pages: [
            {
              ...source.print.displayList.pages[0]!,
              commands: [
                {
                  kind: 'group',
                  rotation: 90,
                  origin: { x: 30, y: 35 },
                  commands: [
                    {
                      kind: 'path',
                      data: 'M 30 20 A 20 15 0 1 0 30 50 A 20 15 0 1 0 30 20 Z',
                      fill: '#ffeecc',
                      stroke: '#112233',
                      width: 2,
                    },
                    {
                      kind: 'clip',
                      rect: { x: 10, y: 20, width: 40, height: 30 },
                      commands: [
                        {
                          kind: 'text',
                          text: 'Clipped',
                          x: 10,
                          y: 20,
                          maxWidth: 40,
                          fontFamily: 'Arial',
                          fontSize: 12,
                          color: '#111111',
                          horizontalAlign: 'left',
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            source.print.displayList.pages[1]!,
          ],
        },
      },
    } as GeneratedDocumentForBrowserPrint;

    const [page] = serializeGeneratedDocumentSvgPages(rotated);

    expect(page?.svg).toContain('<g transform="rotate(90 30 35)">');
    expect(page?.svg).toContain(
      '<path d="M 30 20 A 20 15 0 1 0 30 50 A 20 15 0 1 0 30 20 Z" fill="#ffeecc" stroke="#112233" stroke-width="2"/>',
    );
    expect(page?.svg).toContain('<clipPath id="tego-clip-page-0-0-1">');
    expect(page?.svg.indexOf('transform="rotate')).toBeLessThan(
      page?.svg.indexOf('clip-path=') ?? 0,
    );
  });

  it('serializes a resolved raster resource into both preview and browser-print SVG', () => {
    const source = generatedDocument();
    const documentWithImage = {
      ...source,
      resources: {
        byReference: {
          logo: {
            contentHash: 'sha256:logo',
            type: 'image',
            mimeType: 'image/png',
            bytes: [1, 2],
          },
        },
      },
      print: {
        ...source.print,
        displayList: {
          diagnostics: [],
          pages: [
            {
              ...source.print.displayList.pages[0]!,
              commands: [
                {
                  kind: 'image',
                  resourceId: 'logo',
                  rect: { x: 10, y: 20, width: 40, height: 30 },
                  fit: 'contain',
                },
              ],
            },
            source.print.displayList.pages[1]!,
          ],
        },
      },
    } as GeneratedDocumentForBrowserPrint;

    const [page] = serializeGeneratedDocumentSvgPages(documentWithImage);

    expect(page?.svg).toContain('href="data:image/png;base64,AQI="');
    expect(page?.svg).toContain('preserveAspectRatio="xMidYMid meet"');
    expect(page?.svg).not.toContain('<rect x="10" y="20" width="40" height="30" fill="none"/>');
  });

  it('mounts a hidden same-origin iframe containing only generated SVG pages', async () => {
    const editor = document.createElement('main');
    editor.dataset.editor = '';
    editor.innerHTML = '<button>Editor toolbar</button><aside>Template panel</aside>';
    document.body.append(editor);
    let printed = false;
    let printedPageCount = 0;
    let printedText = '';
    let printedHtml = '';
    let leakedEditorNode: Element | null = null;
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      timeoutMs: 1_000,
      print: (target) => {
        printed = true;
        const mountedFrame = frame();
        expect(target).toBe(mountedFrame.contentWindow);
        expect(mountedFrame.hidden).toBe(true);
        expect(mountedFrame.src).toBe('about:blank');
        printedPageCount = target.document.querySelectorAll('svg[data-page-id]').length;
        printedText = target.document.body.textContent ?? '';
        printedHtml = target.document.body.innerHTML;
        leakedEditorNode = target.document.querySelector('button, aside, [data-editor]');
        target.dispatchEvent(new Event('afterprint'));
      },
    });

    const result = await adapter.print(generatedDocument());

    expect(printed).toBe(true);
    expect(printedPageCount).toBe(2);
    expect(printedText).toContain('Invoice <not-markup>');
    expect(printedHtml).not.toContain('<not-markup>');
    expect(leakedEditorNode).toBeNull();
    expect(document.querySelector('[data-editor]')).toBe(editor);
    expect(result.cleanupReason).toBe('afterprint');
  });

  it('preserves preview and print page identity, count, and geometry', async () => {
    const source = generatedDocument();
    const preview = serializeGeneratedDocumentSvgPages(source);
    let printGeometry: readonly {
      readonly id: string;
      readonly width: string | null;
      readonly height: string | null;
      readonly viewBox: string | null;
    }[] = [];
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      print: (target) => {
        printGeometry = [
          ...target.document.querySelectorAll<SVGSVGElement>('svg[data-page-id]'),
        ].map((page) => ({
          id: page.dataset.pageId ?? '',
          width: page.getAttribute('width'),
          height: page.getAttribute('height'),
          viewBox: page.getAttribute('viewBox'),
        }));
        target.dispatchEvent(new Event('afterprint'));
      },
    });

    const result = await adapter.print(source);

    expect(result.pageIds).toEqual(preview.map((page) => page.id));
    expect(result.pageCount).toBe(preview.length);
    expect(printGeometry).toEqual(
      preview.map((page) => ({
        id: page.id,
        width: String(page.width),
        height: String(page.height),
        viewBox: `0 0 ${page.width} ${page.height}`,
      })),
    );
  });

  it('cleans up after the iframe emits afterprint', async () => {
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      print: (target) => target.dispatchEvent(new Event('afterprint')),
    });

    await expect(adapter.print(generatedDocument())).resolves.toMatchObject({
      cleanupReason: 'afterprint',
    });
    expect(document.querySelector('iframe[data-tego-browser-print]')).toBeNull();
  });

  it('cleans up after a bounded timeout when afterprint is not emitted', async () => {
    vi.useFakeTimers();
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      timeoutMs: 25,
      print: () => undefined,
    });

    const pending = adapter.print(generatedDocument());
    await vi.advanceTimersByTimeAsync(25);

    await expect(pending).resolves.toMatchObject({ cleanupReason: 'timeout' });
    expect(document.querySelector('iframe[data-tego-browser-print]')).toBeNull();
  });

  it('cancels and cleans up through AbortSignal', async () => {
    const controller = new AbortController();
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      print: () => undefined,
    });

    const pending = adapter.print(generatedDocument(), { signal: controller.signal });
    controller.abort();

    await expect(pending).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    expect(document.querySelector('iframe[data-tego-browser-print]')).toBeNull();
  });

  it('rejects an already-aborted request without mounting a frame', async () => {
    const controller = new AbortController();
    controller.abort();
    const adapter = new IsolatedBrowserPrintAdapter({ document });

    await expect(
      adapter.print(generatedDocument(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    expect(document.querySelector('iframe[data-tego-browser-print]')).toBeNull();
  });

  it('dispose is idempotent and aborts every active print session', async () => {
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      print: () => undefined,
    });
    const first = adapter.print(generatedDocument());
    const second = adapter.print(generatedDocument());

    adapter.dispose();
    adapter.dispose();

    await expect(first).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    await expect(second).rejects.toMatchObject({ code: 'RENDER_ABORTED' });
    expect(document.querySelectorAll('iframe[data-tego-browser-print]')).toHaveLength(0);
  });

  it('reports PRINT_BLOCKED and removes the frame when the browser rejects print', async () => {
    const adapter = new IsolatedBrowserPrintAdapter({
      document,
      print: () => {
        throw new DOMException('Blocked by browser policy', 'NotAllowedError');
      },
    });

    await expect(adapter.print(generatedDocument())).rejects.toEqual(
      expect.objectContaining<Partial<BrowserPrintError>>({
        code: 'PRINT_BLOCKED',
        message: expect.stringContaining('blocked'),
      }),
    );
    expect(document.querySelector('iframe[data-tego-browser-print]')).toBeNull();
  });
});
