import type { Diagnostic } from '../../document';

/** Output stage requesting a resolved resource. */
export type ResourcePurpose = 'preview' | 'print' | 'pdf' | 'xlsx' | 'image';
/** Supported logical resource category. */
export type ResourceType = 'image' | 'svg' | 'font' | 'qr' | 'binary';

/** Deterministic QR rendering options. */
export interface QrResourceOptions {
  /** QR error-correction level. */
  readonly errorCorrection?: 'L' | 'M' | 'Q' | 'H';
  /** Foreground CSS color. */
  readonly foreground?: string;
  /** Background CSS color. */
  readonly background?: string;
}

/** Persistent-safe logical reference passed to one explicit resolver. */
export interface ResourceRef {
  /** Logical reference identity. */
  readonly id: string;
  /** Logical resource category. */
  readonly type: ResourceType;
  /** Resolver capability identity. */
  readonly resolverId: string;
  /** Opaque resolver-owned key without credentials. */
  readonly key: string;
  /** Optional declared MIME expectation. */
  readonly expectedMime?: string;
  /** QR-specific deterministic style. */
  readonly qr?: QrResourceOptions;
}

/** Resource resolution and decompression safety budgets. */
export interface ResourceLimits {
  /** Maximum logical references. */
  readonly maxResources: number;
  /** Maximum bytes for one compressed resource. */
  readonly maxResourceBytes: number;
  /** Maximum compressed bytes for the session. */
  readonly maxTotalResourceBytes: number;
  /** Maximum concurrent resolver calls. */
  readonly maxResolveConcurrency: number;
  /** Maximum decoded image pixels. */
  readonly maxPixels: number;
  /** Maximum parsed SVG elements. */
  readonly maxSvgNodes: number;
  /** Maximum resolved fonts. */
  readonly maxFonts: number;
  /** Maximum wall-clock resolution duration. */
  readonly maxResolveTimeMs: number;
  /** Maximum declared decompressed bytes for one resource. */
  readonly maxDecompressedBytes: number;
}

/** Capability-limited context exposed to a resolver. */
export interface ResolveContext {
  /** Session cancellation signal. */
  readonly signal: AbortSignal;
  /** Effective resource safety budgets. */
  readonly limits: ResourceLimits;
  /** Output stage requesting the resource. */
  readonly requestedPurpose: ResourcePurpose;
}

/** Font readiness handle returned by a trusted resolver. */
export interface ResourceFontHandle {
  /** Resolved font-family name. */
  readonly family: string;
  /** Waits until deterministic metrics are ready. */
  readonly waitUntilReady: (signal: AbortSignal) => Promise<void>;
}

/** Untrusted resolver result before MIME, quota and decoder validation. */
export interface UnverifiedResource {
  /** Resolver-provided compressed bytes. */
  readonly bytes: Uint8Array;
  /** Resolver-provided canonical MIME. */
  readonly mimeType: string;
  /** Optional post-decompression byte estimate. */
  readonly decompressedBytes?: number;
  /** Optional decoded width supplied by a trusted decoder. */
  readonly width?: number;
  /** Optional decoded height supplied by a trusted decoder. */
  readonly height?: number;
  /** Optional font readiness handle. */
  readonly font?: ResourceFontHandle;
  /** Idempotent release callback for resolver-owned handles. */
  readonly dispose?: () => void | Promise<void>;
}

/** One host-owned, explicitly registered resource capability. */
export interface ResourceResolver {
  /** Stable resolver identity. */
  readonly id: string;
  /** Reports whether this resolver owns a logical reference. */
  supports(ref: ResourceRef): boolean;
  /** Resolves one reference with restricted context. */
  resolve(ref: ResourceRef, context: ResolveContext): Promise<UnverifiedResource>;
}

/** Ordered immutable collection of explicit resolvers. */
export interface ResourceResolverRegistry {
  /** Resolver snapshot in declaration order. */
  readonly resolvers: readonly ResourceResolver[];
  /** Resolves an exact declared resolver without fallback network access. */
  resolve(ref: ResourceRef): ResourceResolver | undefined;
}

/** Sanitized immutable vector representation. */
export interface ResolvedResourceVector {
  /** Device-independent vector bounds. */
  readonly viewBox: readonly [number, number, number, number];
  /** Restricted SVG path commands. */
  readonly paths: readonly string[];
  /** Foreground CSS color. */
  readonly foreground: string;
  /** Background CSS color. */
  readonly background: string;
}

/** Immutable result returned by a host image decoder. */
export interface DecodedResourceImage {
  /** Decoded pixel width. */
  readonly width: number;
  /** Decoded pixel height. */
  readonly height: number;
  /** Host-owned read-only decoded representation. */
  readonly representation: unknown;
}

/** Validated content-addressed resource shared by output adapters. */
export interface ResolvedResource {
  /** SHA-256 content identity. */
  readonly contentHash: string;
  /** Validated canonical MIME. */
  readonly mimeType: string;
  /** Immutable byte snapshot; adapters create their own typed view when needed. */
  readonly bytes: readonly number[];
  /** Validated pixel width. */
  readonly width?: number;
  /** Validated pixel height. */
  readonly height?: number;
  /** Host decoder's read-only representation. */
  readonly decoded?: unknown;
  /** Sanitized vector representation. */
  readonly vector?: ResolvedResourceVector;
  /** Ready font-family name. */
  readonly fontFamily?: string;
}

/** Session-owned immutable resource mapping with idempotent cleanup. */
export interface ResolvedResourceStore {
  /** Unique resources keyed by content hash. */
  readonly byHash: Readonly<Record<string, ResolvedResource>>;
  /** Logical references mapped to deduplicated resources. */
  readonly byReference: Readonly<Record<string, ResolvedResource>>;
  /** Total fetched compressed bytes. */
  readonly totalBytes: number;
  /** Releases every resolver-owned handle exactly once. */
  dispose(): Promise<void>;
}

/** Explicit inputs for one atomic resolution session. */
export interface ResourcePipelineOptions {
  /** Explicit resolver registry. */
  readonly registry: ResourceResolverRegistry;
  /** Shared render cancellation signal. */
  readonly signal: AbortSignal;
  /** Output stage requesting resources. */
  readonly purpose: ResourcePurpose;
  /** Optional downward or explicit upward budget overrides. */
  readonly limits?: Partial<ResourceLimits>;
  /** Optional host image decoder. */
  readonly decodeImage?: (
    bytes: Uint8Array,
    mimeType: string,
    signal: AbortSignal,
  ) => Promise<DecodedResourceImage>;
}

/** Atomic resource result; a store is absent whenever an error is present. */
export interface ResourceResolutionResult {
  /** Complete ready store. */
  readonly store?: ResolvedResourceStore;
  /** Ordered stable diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
}

/** Byte-budgeted cross-session content cache. */
export interface ResolvedResourceCache {
  /** Reads and promotes one cached resource. */
  get(contentHash: string): ResolvedResource | undefined;
  /** Inserts a resource and evicts least-recently-used entries. */
  put(resource: ResolvedResource, release?: () => void | Promise<void>): Promise<void>;
  /** Releases and removes all cached entries. */
  clear(): Promise<void>;
  /** Current immutable byte footprint. */
  readonly byteLength: number;
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

/** Creates an explicit byte-budgeted LRU cache for host-managed cross-session reuse. */
export function createResolvedResourceCache(maximumBytes: number): ResolvedResourceCache {
  if (!Number.isFinite(maximumBytes) || maximumBytes <= 0) {
    throw new TypeError('Resource cache byte budget must be positive');
  }
  const entries = new Map<
    string,
    { readonly resource: ResolvedResource; readonly release?: () => void | Promise<void> }
  >();
  let bytes = 0;
  const evict = async (): Promise<void> => {
    while (bytes > maximumBytes && entries.size > 0) {
      const oldest = entries.entries().next().value as
        | [
            string,
            {
              readonly resource: ResolvedResource;
              readonly release?: () => void | Promise<void>;
            },
          ]
        | undefined;
      if (oldest === undefined) return;
      entries.delete(oldest[0]);
      bytes -= oldest[1].resource.bytes.length;
      await oldest[1].release?.();
    }
  };
  return {
    get(contentHash) {
      const entry = entries.get(contentHash);
      if (entry === undefined) return undefined;
      entries.delete(contentHash);
      entries.set(contentHash, entry);
      return entry.resource;
    },
    async put(resource, release) {
      const previous = entries.get(resource.contentHash);
      if (previous !== undefined) {
        entries.delete(resource.contentHash);
        bytes -= previous.resource.bytes.length;
        await previous.release?.();
      }
      entries.set(resource.contentHash, {
        resource,
        ...(release === undefined ? {} : { release }),
      });
      bytes += resource.bytes.length;
      await evict();
    },
    async clear() {
      const releases = [...entries.values()].map(({ release }) => release?.());
      entries.clear();
      bytes = 0;
      await Promise.allSettled(releases);
    },
    get byteLength() {
      return bytes;
    },
  };
}

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

function encodedImageDimensions(
  bytes: Uint8Array,
  mimeType: string,
): { readonly width: number; readonly height: number } | undefined {
  if (mimeType === 'image/png' && bytes.length >= 24) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (mimeType === 'image/gif' && bytes.length >= 10) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
  }
  return undefined;
}

function safeSvg(bytes: Uint8Array, maxNodes: number): ResolvedResourceVector | undefined {
  const source = new TextDecoder().decode(bytes);
  const nodes = source.match(/<\s*[a-zA-Z][^>]*>/gu) ?? [];
  if (nodes.length > maxNodes) return undefined;
  if (
    /<\s*(?:script|foreignObject|iframe|object|embed|audio|video)\b/iu.test(source) ||
    /\bon[a-z]+\s*=/iu.test(source) ||
    /\b(?:href|src)\s*=\s*["']\s*(?!#|data:image\/(?:png|jpeg|gif);base64,)/iu.test(source) ||
    /\burl\s*\(\s*["']?\s*(?!#)/iu.test(source) ||
    /<!DOCTYPE|<!ENTITY/iu.test(source)
  ) {
    return undefined;
  }
  const allowedElements = new Set([
    'svg',
    'g',
    'path',
    'rect',
    'circle',
    'ellipse',
    'line',
    'polyline',
    'polygon',
    'defs',
    'clipPath',
    'linearGradient',
    'radialGradient',
    'stop',
    'title',
    'desc',
  ]);
  const tags = [...source.matchAll(/<\s*\/?\s*([a-zA-Z][\w-]*)\b([^>]*)>/gu)];
  if (tags.some((match) => !allowedElements.has(match[1]!))) return undefined;
  const allowedAttributes = new Set([
    'viewBox',
    'd',
    'x',
    'y',
    'x1',
    'x2',
    'y1',
    'y2',
    'width',
    'height',
    'cx',
    'cy',
    'r',
    'rx',
    'ry',
    'points',
    'fill',
    'fill-rule',
    'stroke',
    'stroke-width',
    'stroke-linecap',
    'stroke-linejoin',
    'opacity',
    'transform',
    'id',
    'offset',
    'stop-color',
    'stop-opacity',
    'clip-path',
    'xmlns',
  ]);
  for (const match of tags) {
    const attributes = match[2] ?? '';
    for (const attribute of attributes.matchAll(/\s+([:\w-]+)\s*=/gu)) {
      if (!allowedAttributes.has(attribute[1]!)) return undefined;
    }
  }
  const viewBoxText = /\bviewBox\s*=\s*["']([^"']+)["']/u.exec(source)?.[1];
  const values = viewBoxText?.trim().split(/\s+/u).map(Number);
  const viewBox =
    values?.length === 4 && values.every(Number.isFinite)
      ? (values as [number, number, number, number])
      : ([0, 0, 1, 1] as const);
  const paths = [...source.matchAll(/<\s*path\b[^>]*\bd\s*=\s*["']([^"']*)["']/giu)].map(
    (match) => match[1]!,
  );
  return freeze({
    viewBox,
    paths,
    foreground: '#000000',
    background: 'transparent',
  });
}

function makeQr(
  ref: ResourceRef,
): UnverifiedResource & { readonly vector: ResolvedResource['vector'] } {
  const canonical = JSON.stringify({
    value: ref.key,
    errorCorrection: ref.qr?.errorCorrection ?? 'M',
    foreground: ref.qr?.foreground ?? '#000000',
    background: ref.qr?.background ?? '#ffffff',
  });
  const level = ref.qr?.errorCorrection ?? 'M';
  const configuration = {
    L: { dataCodewords: 19, errorCodewords: 7, formatBits: 1 },
    M: { dataCodewords: 16, errorCodewords: 10, formatBits: 0 },
    Q: { dataCodewords: 13, errorCodewords: 13, formatBits: 3 },
    H: { dataCodewords: 9, errorCodewords: 17, formatBits: 2 },
  }[level];
  const payload = new TextEncoder().encode(ref.key);
  const bits: number[] = [0, 1, 0, 0];
  for (let bit = 7; bit >= 0; bit -= 1) bits.push((payload.length >>> bit) & 1);
  for (const byte of payload) {
    for (let bit = 7; bit >= 0; bit -= 1) bits.push((byte >>> bit) & 1);
  }
  const capacityBits = configuration.dataCodewords * 8;
  if (bits.length > capacityBits) {
    throw new RangeError(`QR payload exceeds version 1 ${level} capacity`);
  }
  bits.push(...Array(Math.min(4, capacityBits - bits.length)).fill(0));
  while (bits.length % 8 !== 0) bits.push(0);
  const dataCodewords: number[] = [];
  for (let offset = 0; offset < bits.length; offset += 8) {
    dataCodewords.push(
      bits.slice(offset, offset + 8).reduce((value, bit) => (value << 1) | bit, 0),
    );
  }
  for (let pad = 0; dataCodewords.length < configuration.dataCodewords; pad += 1) {
    dataCodewords.push(pad % 2 === 0 ? 0xec : 0x11);
  }
  const exponent = Array<number>(512).fill(0);
  const logarithm = Array<number>(256).fill(0);
  let value = 1;
  for (let index = 0; index < 255; index += 1) {
    exponent[index] = value;
    logarithm[value] = index;
    value <<= 1;
    if ((value & 0x100) !== 0) value ^= 0x11d;
  }
  for (let index = 255; index < exponent.length; index += 1) {
    exponent[index] = exponent[index - 255]!;
  }
  const multiply = (left: number, right: number): number =>
    left === 0 || right === 0 ? 0 : exponent[logarithm[left]! + logarithm[right]!]!;
  let generator = [1];
  for (let degree = 0; degree < configuration.errorCodewords; degree += 1) {
    const next = Array<number>(generator.length + 1).fill(0);
    generator.forEach((coefficient, index) => {
      next[index] ^= coefficient;
      next[index + 1] ^= multiply(coefficient, exponent[degree]!);
    });
    generator = next;
  }
  const remainder = Array<number>(configuration.errorCodewords).fill(0);
  for (const byte of dataCodewords) {
    const factor = byte ^ remainder.shift()!;
    remainder.push(0);
    for (let index = 0; index < remainder.length; index += 1) {
      remainder[index] ^= multiply(generator[index + 1]!, factor);
    }
  }
  const codewordBits = [...dataCodewords, ...remainder].flatMap((byte) =>
    Array.from({ length: 8 }, (_, bit) => (byte >>> (7 - bit)) & 1),
  );
  const size = 21;
  const base = Array.from({ length: size }, () => Array<number>(size).fill(-1));
  const set = (x: number, y: number, dark: boolean): void => {
    if (x >= 0 && y >= 0 && x < size && y < size) base[y]![x] = dark ? 1 : 0;
  };
  const finder = (centerX: number, centerY: number): void => {
    for (let deltaY = -4; deltaY <= 4; deltaY += 1) {
      for (let deltaX = -4; deltaX <= 4; deltaX += 1) {
        const distance = Math.max(Math.abs(deltaX), Math.abs(deltaY));
        set(centerX + deltaX, centerY + deltaY, distance !== 2 && distance !== 4);
      }
    }
  };
  finder(3, 3);
  finder(size - 4, 3);
  finder(3, size - 4);
  for (let index = 8; index < size - 8; index += 1) {
    set(6, index, index % 2 === 0);
    set(index, 6, index % 2 === 0);
  }
  const formatCoordinates = [
    ...Array.from({ length: 6 }, (_, index) => [8, index] as const),
    [8, 7] as const,
    [8, 8] as const,
    [7, 8] as const,
    ...Array.from({ length: 6 }, (_, index) => [5 - index, 8] as const),
    ...Array.from({ length: 8 }, (_, index) => [size - 1 - index, 8] as const),
    ...Array.from({ length: 7 }, (_, index) => [8, size - 7 + index] as const),
  ];
  formatCoordinates.forEach(([x, y]) => set(x, y, false));
  set(8, size - 8, true);
  const maskBit = (mask: number, x: number, y: number): boolean =>
    [
      (x + y) % 2 === 0,
      y % 2 === 0,
      x % 3 === 0,
      (x + y) % 3 === 0,
      (Math.floor(y / 2) + Math.floor(x / 3)) % 2 === 0,
      ((x * y) % 2) + ((x * y) % 3) === 0,
      (((x * y) % 2) + ((x * y) % 3)) % 2 === 0,
      (((x + y) % 2) + ((x * y) % 3)) % 2 === 0,
    ][mask]!;
  const candidates = Array.from({ length: 8 }, (_, mask) => {
    const matrix = base.map((row) => [...row]);
    let bitIndex = 0;
    let upward = true;
    for (let right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right -= 1;
      for (let step = 0; step < size; step += 1) {
        const y = upward ? size - 1 - step : step;
        for (const x of [right, right - 1]) {
          if (matrix[y]![x] !== -1) continue;
          const bit = codewordBits[bitIndex++] ?? 0;
          matrix[y]![x] = bit ^ (maskBit(mask, x, y) ? 1 : 0);
        }
      }
      upward = !upward;
    }
    const formatData = (configuration.formatBits << 3) | mask;
    let remainderBits = formatData << 10;
    for (let bit = 14; bit >= 10; bit -= 1) {
      if (((remainderBits >>> bit) & 1) !== 0) remainderBits ^= 0x537 << (bit - 10);
    }
    const format = ((formatData << 10) | remainderBits) ^ 0x5412;
    const first = [
      ...Array.from({ length: 6 }, (_, index) => [8, index] as const),
      [8, 7] as const,
      [8, 8] as const,
      [7, 8] as const,
      ...Array.from({ length: 6 }, (_, index) => [5 - index, 8] as const),
    ];
    const second = [
      ...Array.from({ length: 8 }, (_, index) => [size - 1 - index, 8] as const),
      ...Array.from({ length: 7 }, (_, index) => [8, size - 7 + index] as const),
    ];
    first.forEach(([x, y], index) => {
      matrix[y]![x] = (format >>> index) & 1;
    });
    second.forEach(([x, y], index) => {
      matrix[y]![x] = (format >>> index) & 1;
    });
    matrix[size - 8]![8] = 1;
    let penalty = 0;
    for (const row of matrix) {
      for (let index = 1, run = 1; index < size; index += 1) {
        run = row[index] === row[index - 1] ? run + 1 : 1;
        if (run === 5) penalty += 3;
        else if (run > 5) penalty += 1;
      }
    }
    for (let y = 0; y < size - 1; y += 1) {
      for (let x = 0; x < size - 1; x += 1) {
        const sum =
          matrix[y]![x]! + matrix[y]![x + 1]! + matrix[y + 1]![x]! + matrix[y + 1]![x + 1]!;
        if (sum === 0 || sum === 4) penalty += 3;
      }
    }
    return { matrix, penalty };
  });
  const matrix = candidates.sort((left, right) => left.penalty - right.penalty)[0]!.matrix;
  const paths = matrix.flatMap((row, y) =>
    row.flatMap((dark, x) => (dark === 1 ? [`M${x} ${y}h1v1h-1z`] : [])),
  );
  const vector = freeze({
    viewBox: [0, 0, size, size] as const,
    paths,
    foreground: ref.qr?.foreground ?? '#000000',
    background: ref.qr?.background ?? '#ffffff',
  });
  return { bytes: new TextEncoder().encode(canonical), mimeType: 'image/x-tego-qr', vector };
}

async function controlled<T>(
  task: Promise<T>,
  signal: AbortSignal,
  remainingMs: number,
): Promise<T> {
  if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
  let timeout: ReturnType<typeof setTimeout> | undefined;
  let removeAbort = (): void => {};
  try {
    return await Promise.race([
      task,
      new Promise<T>((_resolve, reject) => {
        const onAbort = (): void => reject(new DOMException('Aborted', 'AbortError'));
        signal.addEventListener('abort', onAbort, { once: true });
        removeAbort = () => signal.removeEventListener('abort', onAbort);
      }),
      new Promise<T>((_resolve, reject) => {
        timeout = setTimeout(
          () => reject(new DOMException('Timed out', 'TimeoutError')),
          Math.max(1, remainingMs),
        );
      }),
    ]);
  } finally {
    removeAbort();
    if (timeout !== undefined) clearTimeout(timeout);
  }
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
    async resolve(ref, context) {
      const match = /^data:([^;,]+)(;base64)?,([\s\S]*)$/u.exec(ref.key);
      if (match === null) throw new Error('Invalid data URL');
      const estimatedBytes =
        match[2] === undefined
          ? match[3]!.length
          : Math.floor((match[3]!.replace(/=+$/u, '').length * 3) / 4);
      if (
        estimatedBytes > context.limits.maxResourceBytes ||
        estimatedBytes > context.limits.maxTotalResourceBytes
      ) {
        throw new RangeError('Data URL exceeds resource limits');
      }
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
    async resolve(ref, context) {
      const blob = blobs.get(ref.key);
      if (blob === undefined) throw new Error('Unknown Blob handle');
      if (
        blob.size > context.limits.maxResourceBytes ||
        blob.size > context.limits.maxTotalResourceBytes
      ) {
        throw new RangeError('Blob exceeds resource limits');
      }
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
          const pendingResource = resolver.resolve(ref, {
            signal: sessionController.signal,
            limits,
            requestedPurpose: options.purpose,
          });
          void pendingResource.then(
            (late) => {
              if (sessionController.signal.aborted) void late.dispose?.();
            },
            () => undefined,
          );
          raw = await controlled(
            pendingResource,
            sessionController.signal,
            limits.maxResolveTimeMs - (Date.now() - started),
          );
        }
      } catch (cause) {
        if (options.signal.aborted) return;
        fail(
          cause instanceof DOMException && cause.name === 'TimeoutError'
            ? 'RESOURCE_TIMEOUT'
            : 'RESOURCE_FETCH_FAILED',
          `Resolver ${ref.resolverId} failed`,
          ref,
        );
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
      const expectedCategory =
        ref.type === 'image'
          ? 'image/'
          : ref.type === 'svg'
            ? 'image/svg+xml'
            : ref.type === 'font'
              ? 'font/'
              : undefined;
      if (
        (ref.expectedMime !== undefined && raw.mimeType !== ref.expectedMime) ||
        (expectedCategory !== undefined &&
          (expectedCategory.endsWith('/')
            ? !raw.mimeType.startsWith(expectedCategory)
            : raw.mimeType !== expectedCategory)) ||
        ((ref.type === 'image' || ref.type === 'svg' || ref.type === 'font') &&
          sniffed === undefined) ||
        (sniffed !== undefined &&
          raw.mimeType !== sniffed &&
          !(raw.mimeType === 'font/otf' && sniffed === 'font/ttf')) ||
        (ref.expectedMime !== undefined && sniffed !== undefined && sniffed !== ref.expectedMime)
      ) {
        fail('RESOURCE_MIME_MISMATCH', `Resource ${ref.id} MIME does not match its content`, ref);
        return;
      }
      const svgVector = ref.type === 'svg' ? safeSvg(raw.bytes, limits.maxSvgNodes) : undefined;
      if (ref.type === 'svg' && svgVector === undefined) {
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
      const encodedDimensions =
        ref.type === 'image' ? encodedImageDimensions(raw.bytes, raw.mimeType) : undefined;
      const declaredWidth = raw.width ?? encodedDimensions?.width;
      const declaredHeight = raw.height ?? encodedDimensions?.height;
      if (
        declaredWidth !== undefined &&
        declaredHeight !== undefined &&
        declaredWidth * declaredHeight > limits.maxPixels
      ) {
        fail('RESOURCE_TOO_LARGE', `Resource ${ref.id} exceeds its pixel quota`, ref);
        return;
      }
      totalBytes += raw.bytes.byteLength;
      const hash = await contentHash(raw.bytes);
      const semanticKey = [
        ref.type,
        raw.mimeType,
        hash,
        raw.font?.family ?? '',
        raw.vector?.foreground ?? svgVector?.foreground ?? '',
        raw.vector?.background ?? svgVector?.background ?? '',
      ].join('\u0000');
      let resolution = decodeByHash.get(semanticKey);
      if (resolution === undefined) {
        resolution = (async (): Promise<ResolvedResource> => {
          let decoded: unknown;
          let width = declaredWidth;
          let height = declaredHeight;
          if (ref.type === 'image' && options.decodeImage !== undefined) {
            const image = await controlled(
              options.decodeImage(raw.bytes, raw.mimeType, sessionController.signal),
              sessionController.signal,
              limits.maxResolveTimeMs - (Date.now() - started),
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
          if (raw.font !== undefined)
            await controlled(
              raw.font.waitUntilReady(sessionController.signal),
              sessionController.signal,
              limits.maxResolveTimeMs - (Date.now() - started),
            );
          return freeze({
            contentHash: hash,
            mimeType: raw.mimeType,
            bytes: Object.freeze([...raw.bytes]),
            ...(width === undefined ? {} : { width }),
            ...(height === undefined ? {} : { height }),
            ...(decoded === undefined ? {} : { decoded }),
            ...(raw.vector === undefined && svgVector === undefined
              ? {}
              : { vector: raw.vector ?? svgVector }),
            ...(raw.font === undefined ? {} : { fontFamily: raw.font.family }),
          });
        })();
        decodeByHash.set(semanticKey, resolution);
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
