import { describe, expect, it, vi } from 'vitest';
import {
  createResourceResolverRegistry,
  createResolvedResourceCache,
  resolveTemplateResources,
  type ResourceRef,
  type ResourceResolver,
} from '../../../src/template';

const bytes = (...values: number[]) => new Uint8Array(values);

function resolver(id: string, resolve: ResourceResolver['resolve']): ResourceResolver {
  return { id, supports: (ref) => ref.resolverId === id, resolve };
}

describe('TP3 resource pipeline', () => {
  it('deduplicates equal content hashes and decodes one time', async () => {
    const decode = vi.fn(async () => ({ width: 1, height: 1, representation: 'pixel' }));
    const registry = createResourceResolverRegistry([
      resolver('app', async (ref) => ({
        bytes: bytes(1, 2, 3),
        mimeType: ref.expectedMime ?? 'image/png',
      })),
    ]);
    const refs: ResourceRef[] = [
      { id: 'a', type: 'image', resolverId: 'app', key: 'one', expectedMime: 'image/png' },
      { id: 'b', type: 'image', resolverId: 'app', key: 'two', expectedMime: 'image/png' },
    ];
    const result = await resolveTemplateResources(refs, {
      registry,
      signal: new AbortController().signal,
      purpose: 'preview',
      limits: {
        maxResources: 3,
        maxResourceBytes: 10,
        maxTotalResourceBytes: 20,
        maxResolveConcurrency: 2,
        maxPixels: 10,
        maxSvgNodes: 10,
        maxFonts: 2,
        maxResolveTimeMs: 1_000,
        maxDecompressedBytes: 20,
      },
      decodeImage: decode,
    });
    expect(result.store?.byReference.a).toBe(result.store?.byReference.b);
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('rejects forged MIME, oversized/decompression-bomb content and unsafe SVG', async () => {
    const registry = createResourceResolverRegistry([
      resolver('bad', async (ref) => ({
        bytes:
          ref.key === 'svg'
            ? new TextEncoder().encode('<svg><script>alert(1)</script></svg>')
            : bytes(0x89, 0x50, 0x4e, 0x47),
        mimeType:
          ref.key === 'mime' ? 'text/plain' : ref.key === 'svg' ? 'image/svg+xml' : 'image/png',
        decompressedBytes: ref.key === 'bomb' ? 10_000 : undefined,
      })),
    ]);
    const limits = {
      maxResources: 4,
      maxResourceBytes: 100,
      maxTotalResourceBytes: 200,
      maxResolveConcurrency: 1,
      maxPixels: 10,
      maxSvgNodes: 10,
      maxFonts: 1,
      maxResolveTimeMs: 1_000,
      maxDecompressedBytes: 100,
    };
    for (const [key, expected] of [
      ['mime', 'RESOURCE_MIME_MISMATCH'],
      ['svg', 'UNSAFE_SVG'],
      ['bomb', 'RESOURCE_TOO_LARGE'],
    ] as const) {
      const result = await resolveTemplateResources(
        [
          {
            id: key,
            type: key === 'svg' ? 'svg' : 'image',
            resolverId: 'bad',
            key,
            expectedMime: key === 'svg' ? 'image/svg+xml' : 'image/png',
          },
        ],
        { registry, signal: new AbortController().signal, purpose: 'print', limits },
      );
      expect(result.store).toBeUndefined();
      expect(result.diagnostics).toContainEqual(expect.objectContaining({ code: expected }));
    }
  });

  it('never performs implicit network resolution', async () => {
    const registry = createResourceResolverRegistry([]);
    const result = await resolveTemplateResources(
      [
        {
          id: 'remote',
          type: 'image',
          resolverId: 'network',
          key: 'https://example.invalid/a.png',
        },
      ],
      { registry, signal: new AbortController().signal, purpose: 'preview' },
    );
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'RESOURCE_RESOLVER_NOT_FOUND' }),
    );
  });

  it('bounds resolvers that ignore cancellation with a stable timeout', async () => {
    const result = await resolveTemplateResources(
      [{ id: 'hung', type: 'binary', resolverId: 'hung', key: 'hung' }],
      {
        registry: createResourceResolverRegistry([
          resolver('hung', async () => new Promise(() => {})),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
        limits: {
          maxResources: 1,
          maxResourceBytes: 10,
          maxTotalResourceBytes: 10,
          maxResolveConcurrency: 1,
          maxPixels: 1,
          maxSvgNodes: 1,
          maxFonts: 1,
          maxResolveTimeMs: 5,
          maxDecompressedBytes: 10,
        },
      },
    );
    expect(result.store).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'RESOURCE_TIMEOUT' }),
    );
  });

  it('waits for fonts, creates deterministic QR vectors, and disposes on cancellation', async () => {
    const dispose = vi.fn();
    const waitUntilReady = vi.fn(async () => {});
    const registry = createResourceResolverRegistry([
      resolver('font', async () => ({
        bytes: bytes(0, 1, 0, 0),
        mimeType: 'font/ttf',
        font: { family: 'Noto Sans CJK', waitUntilReady },
        dispose,
      })),
    ]);
    const result = await resolveTemplateResources(
      [
        { id: 'font', type: 'font', resolverId: 'font', key: 'cjk', expectedMime: 'font/ttf' },
        {
          id: 'qr',
          type: 'qr',
          resolverId: 'core:qr',
          key: 'invoice-42',
          qr: { errorCorrection: 'M', foreground: '#000000', background: '#ffffff' },
        },
      ],
      { registry, signal: new AbortController().signal, purpose: 'pdf' },
    );
    expect(waitUntilReady).toHaveBeenCalledOnce();
    expect(result.store?.byReference.qr?.vector?.paths.length).toBeGreaterThan(0);
    await result.store?.dispose();
    await result.store?.dispose();
    expect(dispose).toHaveBeenCalledOnce();

    const controller = new AbortController();
    const pending = resolveTemplateResources(
      [{ id: 'slow', type: 'image', resolverId: 'slow', key: 'slow' }],
      {
        registry: createResourceResolverRegistry([
          resolver('slow', async (_ref, context) => {
            await new Promise<void>((resolve) =>
              context.signal.addEventListener('abort', () => resolve(), { once: true }),
            );
            return { bytes: bytes(1), mimeType: 'image/png', dispose };
          }),
        ]),
        signal: controller.signal,
        purpose: 'preview',
      },
    );
    controller.abort();
    const aborted = await pending;
    expect(aborted.store).toBeUndefined();
    expect(aborted.diagnostics).toContainEqual(expect.objectContaining({ code: 'RENDER_ABORTED' }));
  });

  it('keeps concurrent sessions isolated and evicts cache entries by LRU byte budget', async () => {
    let generation = 0;
    const registry = createResourceResolverRegistry([
      resolver('session', async () => ({
        bytes: bytes(1, ++generation),
        mimeType: 'application/octet-stream',
      })),
    ]);
    const [left, right] = await Promise.all([
      resolveTemplateResources(
        [{ id: 'value', type: 'binary', resolverId: 'session', key: 'value' }],
        { registry, signal: new AbortController().signal, purpose: 'preview' },
      ),
      resolveTemplateResources(
        [{ id: 'value', type: 'binary', resolverId: 'session', key: 'value' }],
        { registry, signal: new AbortController().signal, purpose: 'preview' },
      ),
    ]);
    expect(left.store).not.toBe(right.store);
    expect(left.store?.byReference.value).not.toBe(right.store?.byReference.value);

    const releaseFirst = vi.fn();
    const cache = createResolvedResourceCache(3);
    const first = left.store!.byReference.value!;
    const second = right.store!.byReference.value!;
    await cache.put(first, releaseFirst);
    expect(cache.get(first.contentHash)).toBe(first);
    await cache.put({ ...second, bytes: bytes(1, 2, 3, 4) });
    expect(releaseFirst).toHaveBeenCalledOnce();
    expect(cache.byteLength).toBe(0);
  });
});
