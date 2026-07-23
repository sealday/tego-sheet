import { describe, expect, it, vi } from 'vitest';
import {
  createResourceResolverRegistry,
  createResolvedResourceCache,
  resolveTemplateResources,
  type ResourceRef,
  type ResourceResolver,
} from '../../../src/template';

const bytes = (...values: number[]) => new Uint8Array(values);

function png(width = 1, height = 1): Uint8Array {
  const value = new Uint8Array(45);
  value.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  new DataView(value.buffer).setUint32(8, 13);
  value.set(new TextEncoder().encode('IHDR'), 12);
  new DataView(value.buffer).setUint32(16, width);
  new DataView(value.buffer).setUint32(20, height);
  new DataView(value.buffer).setUint32(33, 0);
  value.set(new TextEncoder().encode('IEND'), 37);
  return value;
}

function resolver(id: string, resolve: ResourceResolver['resolve']): ResourceResolver {
  return { id, supports: (ref) => ref.resolverId === id, resolve };
}

describe('TP3 resource pipeline', () => {
  it('deduplicates equal content hashes and decodes one time', async () => {
    const decode = vi.fn(async () => ({ width: 1, height: 1, representation: 'pixel' }));
    const registry = createResourceResolverRegistry([
      resolver('app', async (ref) => ({
        bytes: png(),
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
        maxResourceBytes: 100,
        maxTotalResourceBytes: 200,
        maxResolveConcurrency: 2,
        maxPixels: 10,
        maxSvgNodes: 10,
        maxFonts: 2,
        maxResolveTimeMs: 1_000,
        maxDecompressedBytes: 200,
      },
      decodeImage: decode,
    });
    expect(result.store?.byReference.a).toBe(result.store?.byReference.b);
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('rejects a raster image that has a signature but no complete image structure', async () => {
    const result = await resolveTemplateResources(
      [{ id: 'truncated', type: 'image', resolverId: 'app', key: 'truncated' }],
      {
        registry: createResourceResolverRegistry([
          resolver('app', async () => ({
            bytes: bytes(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a),
            mimeType: 'image/png',
          })),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
      },
    );

    expect(result.store).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'RESOURCE_DECODE_FAILED' }),
    );
  });

  it('requires a real raster decoder instead of accepting structural marker bytes', async () => {
    const result = await resolveTemplateResources(
      [{ id: 'forged', type: 'image', resolverId: 'app', key: 'forged' }],
      {
        registry: createResourceResolverRegistry([
          resolver('app', async () => ({ bytes: png(), mimeType: 'image/png' })),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
      },
    );

    expect(result.store).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'RESOURCE_DECODE_FAILED' }),
    );
  });

  it('disposes an image decoder result that arrives after the decode timeout', async () => {
    const lateDispose = vi.fn();
    const result = await resolveTemplateResources(
      [{ id: 'late-image', type: 'image', resolverId: 'app', key: 'late-image' }],
      {
        registry: createResourceResolverRegistry([
          resolver('app', async () => ({ bytes: png(), mimeType: 'image/png' })),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
        limits: {
          maxResources: 1,
          maxResourceBytes: 100,
          maxTotalResourceBytes: 100,
          maxResolveConcurrency: 1,
          maxPixels: 10,
          maxSvgNodes: 10,
          maxFonts: 1,
          maxResolveTimeMs: 5,
          maxDecompressedBytes: 100,
        },
        decodeImage: async () =>
          new Promise((resolve) =>
            setTimeout(
              () => resolve({ width: 1, height: 1, representation: {}, dispose: lateDispose }),
              20,
            ),
          ),
      },
    );

    expect(result.store).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'RESOURCE_DECODE_FAILED' }),
    );
    await new Promise((resolve) => setTimeout(resolve, 30));
    expect(lateDispose).toHaveBeenCalledOnce();
  });

  it('snapshots resolver bytes before decoding and disposes decoded handles with the session', async () => {
    const decodedDispose = vi.fn();
    const source = png();
    const result = await resolveTemplateResources(
      [{ id: 'logo', type: 'image', resolverId: 'app', key: 'logo' }],
      {
        registry: createResourceResolverRegistry([
          resolver('app', async () => ({ bytes: source, mimeType: 'image/png' })),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
        decodeImage: async (input) => {
          input[16] = 0xff;
          return { width: 1, height: 1, representation: {}, dispose: decodedDispose };
        },
      },
    );

    const stored = result.store?.byReference.logo;
    expect(stored).toBeDefined();
    const digest = await crypto.subtle.digest('SHA-256', Uint8Array.from(stored!.bytes));
    const actual = `sha256:${[...new Uint8Array(digest)]
      .map((value) => value.toString(16).padStart(2, '0'))
      .join('')}`;
    expect(stored?.contentHash).toBe(actual);
    expect(stored?.bytes[16]).toBe(0);
    await result.store?.dispose();
    await result.store?.dispose();
    expect(decodedDispose).toHaveBeenCalledOnce();
  });

  it.each([
    ['font/woff', [0x77, 0x4f, 0x46, 0x46]],
    ['font/woff2', [0x77, 0x4f, 0x46, 0x32]],
  ] as const)('accepts the canonical %s MIME for matching font bytes', async (mimeType, prefix) => {
    const result = await resolveTemplateResources(
      [{ id: mimeType, type: 'font', resolverId: 'font', key: mimeType, expectedMime: mimeType }],
      {
        registry: createResourceResolverRegistry([
          resolver('font', async () => ({
            bytes: bytes(...prefix),
            mimeType,
            font: { family: 'Test', waitUntilReady: async () => {} },
          })),
        ]),
        signal: new AbortController().signal,
        purpose: 'preview',
      },
    );

    expect(result.diagnostics).toEqual([]);
    expect(result.store?.byReference[mimeType]?.mimeType).toBe(mimeType);
  });

  it('rejects duplicate logical resource IDs before invoking a resolver', async () => {
    const resolve = vi.fn(async () => ({
      bytes: bytes(1),
      mimeType: 'application/octet-stream',
    }));
    const result = await resolveTemplateResources(
      [
        { id: 'same', type: 'binary', resolverId: 'app', key: 'first' },
        { id: 'same', type: 'binary', resolverId: 'app', key: 'second' },
      ],
      {
        registry: createResourceResolverRegistry([resolver('app', resolve)]),
        signal: new AbortController().signal,
        purpose: 'preview',
      },
    );

    expect(result.store).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'DUPLICATE_RESOURCE_ID' }),
    );
    expect(resolve).not.toHaveBeenCalled();
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
    const lateDispose = vi.fn();
    const result = await resolveTemplateResources(
      [{ id: 'hung', type: 'binary', resolverId: 'hung', key: 'hung' }],
      {
        registry: createResourceResolverRegistry([
          resolver(
            'hung',
            async () =>
              new Promise((resolve) =>
                setTimeout(
                  () =>
                    resolve({
                      bytes: bytes(1),
                      mimeType: 'application/octet-stream',
                      dispose: lateDispose,
                    }),
                  20,
                ),
              ),
          ),
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
    await new Promise((resolve) => setTimeout(resolve, 30));
    expect(lateDispose).toHaveBeenCalledOnce();
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
    await cache.put({ ...second, bytes: [1, 2, 3, 4] });
    expect(releaseFirst).toHaveBeenCalledOnce();
    expect(cache.byteLength).toBe(0);
  });
});
