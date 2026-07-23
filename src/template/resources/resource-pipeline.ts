import type { Diagnostic } from '../../document';

export type ResourcePurpose = 'preview' | 'print' | 'pdf' | 'xlsx' | 'image';
export type ResourceType = 'image' | 'svg' | 'font' | 'qr' | 'binary';

export interface ResourceRef {
  readonly id: string;
  readonly type: ResourceType;
  readonly resolverId: string;
  readonly key: string;
  readonly expectedMime?: string;
  readonly qr?: {
    readonly errorCorrection?: 'L' | 'M' | 'Q' | 'H';
    readonly foreground?: string;
    readonly background?: string;
  };
}

export interface ResourceLimits {
  readonly maxResources: number;
  readonly maxResourceBytes: number;
  readonly maxTotalResourceBytes: number;
  readonly maxResolveConcurrency: number;
  readonly maxPixels: number;
  readonly maxSvgNodes: number;
  readonly maxFonts: number;
  readonly maxResolveTimeMs: number;
  readonly maxDecompressedBytes: number;
}

export interface ResolveContext {
  readonly signal: AbortSignal;
  readonly limits: ResourceLimits;
  readonly requestedPurpose: ResourcePurpose;
}

export interface UnverifiedResource {
  readonly bytes: Uint8Array;
  readonly mimeType: string;
  readonly decompressedBytes?: number;
  readonly width?: number;
  readonly height?: number;
  readonly font?: {
    readonly family: string;
    readonly waitUntilReady: (signal: AbortSignal) => Promise<void>;
  };
  readonly dispose?: () => void | Promise<void>;
}

export interface ResourceResolver {
  readonly id: string;
  supports(ref: ResourceRef): boolean;
  resolve(ref: ResourceRef, context: ResolveContext): Promise<UnverifiedResource>;
}

export interface ResourceResolverRegistry {
  readonly resolvers: readonly ResourceResolver[];
  resolve(ref: ResourceRef): ResourceResolver | undefined;
}

export interface ResolvedResource {
  readonly contentHash: string;
  readonly mimeType: string;
  readonly bytes: Uint8Array;
  readonly width?: number;
  readonly height?: number;
  readonly decoded?: unknown;
  readonly vector?: {
    readonly viewBox: readonly [number, number, number, number];
    readonly paths: readonly string[];
    readonly foreground: string;
    readonly background: string;
  };
  readonly fontFamily?: string;
}

export interface ResolvedResourceStore {
  readonly byHash: Readonly<Record<string, ResolvedResource>>;
  readonly byReference: Readonly<Record<string, ResolvedResource>>;
  readonly totalBytes: number;
  dispose(): Promise<void>;
}

export interface ResourcePipelineOptions {
  readonly registry: ResourceResolverRegistry;
  readonly signal: AbortSignal;
  readonly purpose: ResourcePurpose;
  readonly limits?: Partial<ResourceLimits>;
  readonly decodeImage?: (
    bytes: Uint8Array,
    mimeType: string,
    signal: AbortSignal,
  ) => Promise<{
    readonly width: number;
    readonly height: number;
    readonly representation: unknown;
  }>;
}

export interface ResourceResolutionResult {
  readonly store?: ResolvedResourceStore;
  readonly diagnostics: readonly Diagnostic[];
}

const DEFAULT_LIMITS: ResourceLimits = Object.freeze({
  maxResources: 256,
  maxResourceBytes: 16 * 1024 * 1024,
  maxTotalResourceBytes: 64 * 1024 * 1024,
  maxResolveConcurrency: 4,
  maxPixels: 40_000_000,
  maxSvgNodes: 20_000,
  maxFonts: 16,
  maxResolveTimeMs: 15_000,
  maxDecompressedBytes: 64 * 1024 * 1024,
});

function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (value !== null && typeof value === 'object' && !(value instanceof Uint8Array)) {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

function diagnostic(code: string, message: string, resourceId?: string): Diagnostic {
  return freeze({
    code,
    severity: 'error',
    domain: 'template',
    stage: 'resolve',
    message,
    ...(resourceId === undefined ? {} : { location: { resourceId: resourceId as never } }),
  });
}

function aborted(): ResourceResolutionResult {
  return freeze({
    diagnostics: [diagnostic('RENDER_ABORTED', 'Resource resolution was aborted')],
  });
}

function mergeLimits(limits?: Partial<ResourceLimits>): ResourceLimits {
  return Object.freeze({ ...DEFAULT_LIMITS, ...limits });
}

function validLimits(limits: ResourceLimits): boolean {
  return Object.values(limits).every((value) => Number.isFinite(value) && value > 0);
}

async function contentHash(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', Uint8Array.from(bytes));
  return `sha256:${[...new Uint8Array(digest)]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

function sniffMime(bytes: Uint8Array): string | undefined {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return 'image/png';
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg';
  if (
    bytes.length >= 6 &&
    new TextDecoder().decode(bytes.slice(0, 6)).toUpperCase().startsWith('GIF')
  ) {
    return 'image/gif';
  }
  const prefix = new TextDecoder().decode(bytes.slice(0, Math.min(bytes.length, 256))).trimStart();
  if (prefix.startsWith('<svg') || prefix.startsWith('<?xml')) return 'image/svg+xml';
  if (
    (bytes[0] === 0x00 && bytes[1] === 0x01 && bytes[2] === 0x00 && bytes[3] === 0x00) ||
    new TextDecoder().decode(bytes.slice(0, 4)) === 'OTTO' ||
    new TextDecoder().decode(bytes.slice(0, 4)) === 'wOFF' ||
    new TextDecoder().decode(bytes.slice(0, 4)) === 'wOF2'
  ) {
    return 'font/ttf';
  }
  return undefined;
}

function safeSvg(bytes: Uint8Array, maxNodes: number): boolean {
  const source = new TextDecoder().decode(bytes);
  const nodes = source.match(/<\s*[a-zA-Z][^>]*>/gu) ?? [];
  if (nodes.length > maxNodes) return false;
  return !(
    /<\s*(?:script|foreignObject|iframe|object|embed|audio|video)\b/iu.test(source) ||
    /\bon[a-z]+\s*=/iu.test(source) ||
    /\b(?:href|src)\s*=\s*["']\s*(?!#|data:image\/(?:png|jpeg|gif);base64,)/iu.test(source) ||
    /\burl\s*\(\s*["']?\s*(?!#)/iu.test(source) ||
    /<!DOCTYPE|<!ENTITY/iu.test(source)
  );
}

function makeQr(
  ref: ResourceRef,
): UnverifiedResource & { readonly vector: ResolvedResource['vector'] } {
  const size = 21;
  let state = 2166136261;
  for (const value of new TextEncoder().encode(ref.key)) {
    state ^= value;
    state = Math.imul(state, 16777619) >>> 0;
  }
  const occupied = new Set<string>();
  const paths: string[] = [];
  const module = (x: number, y: number): void => {
    const key = `${x}:${y}`;
    if (occupied.has(key)) return;
    occupied.add(key);
    paths.push(`M${x} ${y}h1v1h-1z`);
  };
  for (const [originX, originY] of [
    [0, 0],
    [14, 0],
    [0, 14],
  ] as const) {
    for (let y = 0; y < 7; y += 1) {
      for (let x = 0; x < 7; x += 1) {
        if (x === 0 || y === 0 || x === 6 || y === 6 || (x >= 2 && x <= 4 && y >= 2 && y <= 4)) {
          module(originX + x, originY + y);
        }
      }
    }
  }
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
      if ((state & 3) === 0) module(x, y);
    }
  }
  const vector = freeze({
    viewBox: [0, 0, size, size] as const,
    paths,
    foreground: ref.qr?.foreground ?? '#000000',
    background: ref.qr?.background ?? '#ffffff',
  });
  return { bytes: new TextEncoder().encode(ref.key), mimeType: 'image/x-tego-qr', vector };
}

/** Creates a deterministic, explicit resolver registry. No network resolver is installed implicitly. */
export function createResourceResolverRegistry(
  resolvers: readonly ResourceResolver[],
): ResourceResolverRegistry {
  const seen = new Set<string>();
  const snapshot = resolvers.map((resolver) => {
    if (seen.has(resolver.id)) throw new Error(`Duplicate resource resolver ${resolver.id}`);
    seen.add(resolver.id);
    return resolver;
  });
  return Object.freeze({
    resolvers: Object.freeze(snapshot),
    resolve(ref: ResourceRef): ResourceResolver | undefined {
      return snapshot.find((resolver) => resolver.id === ref.resolverId && resolver.supports(ref));
    },
  });
}

/** Resolver for explicit `data:` values. */
export function createDataUrlResourceResolver(): ResourceResolver {
  return {
    id: 'core:data-url',
    supports: ({ resolverId }) => resolverId === 'core:data-url',
    async resolve(ref) {
      const match = /^data:([^;,]+)(;base64)?,([\s\S]*)$/u.exec(ref.key);
      if (match === null) throw new Error('Invalid data URL');
      const bytes =
        match[2] === undefined
          ? new TextEncoder().encode(decodeURIComponent(match[3]!))
          : Uint8Array.from(atob(match[3]!), (character) => character.charCodeAt(0));
      return { bytes, mimeType: match[1]! };
    },
  };
}

/** Resolver for application-owned Blob handles; blobs are never persisted in the document. */
export function createBlobResourceResolver(blobs: ReadonlyMap<string, Blob>): ResourceResolver {
  return {
    id: 'core:blob',
    supports: ({ resolverId, key }) => resolverId === 'core:blob' && blobs.has(key),
    async resolve(ref) {
      const blob = blobs.get(ref.key);
      if (blob === undefined) throw new Error('Unknown Blob handle');
      return {
        bytes: new Uint8Array(await blob.arrayBuffer()),
        mimeType: blob.type || 'application/octet-stream',
      };
    },
  };
}

/** Resolves, validates, hashes, deduplicates and readies one atomic resource session. */
export async function resolveTemplateResources(
  refs: readonly ResourceRef[],
  options: ResourcePipelineOptions,
): Promise<ResourceResolutionResult> {
  const limits = mergeLimits(options.limits);
  if (!validLimits(limits) || refs.length > limits.maxResources) {
    return freeze({
      diagnostics: [
        diagnostic('RESOURCE_TOO_LARGE', `Resource count exceeds ${limits.maxResources}`),
      ],
    });
  }
  if (options.signal.aborted) return aborted();
  const started = Date.now();
  const sessionController = new AbortController();
  const abort = (): void => sessionController.abort();
  options.signal.addEventListener('abort', abort, { once: true });
  const disposers = new Set<() => void | Promise<void>>();
  const byHash: Record<string, ResolvedResource> = Object.create(null);
  const byReference: Record<string, ResolvedResource> = Object.create(null);
  const diagnostics: Diagnostic[] = [];
  const decodeByHash = new Map<string, Promise<ResolvedResource>>();
  let totalBytes = 0;
  let fonts = 0;
  let next = 0;
  let disposed = false;
  const dispose = async (): Promise<void> => {
    if (disposed) return;
    disposed = true;
    sessionController.abort();
    await Promise.allSettled([...disposers].map((release) => release()));
    disposers.clear();
  };
  const fail = (code: string, message: string, ref: ResourceRef): void => {
    diagnostics.push(diagnostic(code, message, ref.id));
    sessionController.abort();
  };
  const worker = async (): Promise<void> => {
    while (!sessionController.signal.aborted) {
      const index = next++;
      if (index >= refs.length) return;
      const ref = refs[index]!;
      if (Date.now() - started > limits.maxResolveTimeMs) {
        fail('RESOURCE_TIMEOUT', 'Resource resolution exceeded its time limit', ref);
        return;
      }
      let raw: UnverifiedResource & { readonly vector?: ResolvedResource['vector'] };
      try {
        if (ref.type === 'qr' && ref.resolverId === 'core:qr') {
          raw = makeQr(ref);
        } else {
          const resolver = options.registry.resolve(ref);
          if (resolver === undefined) {
            fail(
              'RESOURCE_RESOLVER_NOT_FOUND',
              `No resolver registered for ${ref.resolverId}`,
              ref,
            );
            return;
          }
          raw = await resolver.resolve(ref, {
            signal: sessionController.signal,
            limits,
            requestedPurpose: options.purpose,
          });
        }
      } catch {
        if (options.signal.aborted) return;
        fail('RESOURCE_FETCH_FAILED', `Resolver ${ref.resolverId} failed`, ref);
        return;
      }
      if (raw.dispose !== undefined) disposers.add(raw.dispose);
      if (sessionController.signal.aborted) return;
      const decompressed = raw.decompressedBytes ?? raw.bytes.byteLength;
      if (
        raw.bytes.byteLength > limits.maxResourceBytes ||
        decompressed > limits.maxDecompressedBytes ||
        totalBytes + raw.bytes.byteLength > limits.maxTotalResourceBytes
      ) {
        fail('RESOURCE_TOO_LARGE', `Resource ${ref.id} exceeds its byte quota`, ref);
        return;
      }
      const sniffed = sniffMime(raw.bytes);
      if (
        (ref.expectedMime !== undefined && raw.mimeType !== ref.expectedMime) ||
        (sniffed !== undefined &&
          raw.mimeType !== sniffed &&
          !(raw.mimeType === 'font/otf' && sniffed === 'font/ttf')) ||
        (ref.expectedMime !== undefined && sniffed !== undefined && sniffed !== ref.expectedMime)
      ) {
        fail('RESOURCE_MIME_MISMATCH', `Resource ${ref.id} MIME does not match its content`, ref);
        return;
      }
      if (ref.type === 'svg' && !safeSvg(raw.bytes, limits.maxSvgNodes)) {
        fail('UNSAFE_SVG', `Resource ${ref.id} contains unsafe SVG content`, ref);
        return;
      }
      if (ref.type === 'font') {
        fonts += 1;
        if (fonts > limits.maxFonts || raw.font === undefined || sniffed === undefined) {
          fail('FONT_PARSE_FAILED', `Font ${ref.id} could not be parsed`, ref);
          return;
        }
      }
      totalBytes += raw.bytes.byteLength;
      const hash = await contentHash(raw.bytes);
      let resolution = decodeByHash.get(hash);
      if (resolution === undefined) {
        resolution = (async (): Promise<ResolvedResource> => {
          let decoded: unknown;
          let width = raw.width;
          let height = raw.height;
          if (ref.type === 'image' && options.decodeImage !== undefined) {
            const image = await options.decodeImage(
              raw.bytes,
              raw.mimeType,
              sessionController.signal,
            );
            width = image.width;
            height = image.height;
            decoded = image.representation;
          }
          if (
            width !== undefined &&
            height !== undefined &&
            (!Number.isFinite(width) ||
              !Number.isFinite(height) ||
              width <= 0 ||
              height <= 0 ||
              width * height > limits.maxPixels)
          ) {
            throw new Error('PIXEL_LIMIT');
          }
          if (raw.font !== undefined) await raw.font.waitUntilReady(sessionController.signal);
          return freeze({
            contentHash: hash,
            mimeType: raw.mimeType,
            bytes: raw.bytes.slice(),
            ...(width === undefined ? {} : { width }),
            ...(height === undefined ? {} : { height }),
            ...(decoded === undefined ? {} : { decoded }),
            ...(raw.vector === undefined ? {} : { vector: raw.vector }),
            ...(raw.font === undefined ? {} : { fontFamily: raw.font.family }),
          });
        })();
        decodeByHash.set(hash, resolution);
      }
      try {
        const resolved = await resolution!;
        byHash[hash] = resolved;
        byReference[ref.id] = resolved;
      } catch {
        if (!options.signal.aborted) {
          fail(
            ref.type === 'font' ? 'FONT_PARSE_FAILED' : 'RESOURCE_DECODE_FAILED',
            `Resource ${ref.id} could not be decoded`,
            ref,
          );
        }
        return;
      }
    }
  };
  try {
    await Promise.all(
      Array.from(
        { length: Math.min(refs.length, Math.max(1, limits.maxResolveConcurrency)) },
        worker,
      ),
    );
  } finally {
    options.signal.removeEventListener('abort', abort);
  }
  if (options.signal.aborted) {
    await dispose();
    return aborted();
  }
  if (diagnostics.length > 0 || sessionController.signal.aborted) {
    await dispose();
    return freeze({ diagnostics });
  }
  const store: ResolvedResourceStore = freeze({
    byHash,
    byReference,
    totalBytes,
    dispose,
  });
  return freeze({ store, diagnostics: [] });
}
