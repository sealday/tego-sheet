import type { CellPresentation } from './cell-presentation';

/** Explicit memory limits for one presentation cache. */
export interface PresentationCacheOptions {
  /** Maximum number of retained presentations. */
  readonly maximumEntries: number;
  /** Maximum approximate serialized bytes retained. */
  readonly maximumBytes: number;
}

/** Observable bounded-cache counters. */
export interface PresentationCacheStats {
  /** Current retained entry count. */
  readonly entries: number;
  /** Current approximate retained byte count. */
  readonly bytes: number;
  /** Successful lookup count. */
  readonly hits: number;
  /** Failed lookup count. */
  readonly misses: number;
  /** Budget-driven eviction count. */
  readonly evictions: number;
}

/** Bounded LRU storage for immutable cell presentations. */
export interface PresentationCache {
  /** Reads and refreshes one LRU entry. */
  get(key: string): CellPresentation | undefined;
  /** Inserts one immutable presentation and enforces budgets. */
  set(key: string, value: CellPresentation): void;
  /** Removes every retained entry. */
  clear(): void;
  /** Returns immutable cache counters. */
  stats(): PresentationCacheStats;
}

function byteSize(value: CellPresentation): number {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

/** Creates an isolated presentation cache with entry and byte budgets. */
export function createPresentationCache(options: PresentationCacheOptions): PresentationCache {
  if (!Number.isSafeInteger(options.maximumEntries) || options.maximumEntries <= 0) {
    throw new RangeError('maximumEntries must be a positive safe integer');
  }
  if (!Number.isSafeInteger(options.maximumBytes) || options.maximumBytes <= 0) {
    throw new RangeError('maximumBytes must be a positive safe integer');
  }
  const entries = new Map<string, { readonly value: CellPresentation; readonly bytes: number }>();
  let bytes = 0;
  let hits = 0;
  let misses = 0;
  let evictions = 0;

  const evict = (): void => {
    while (entries.size > options.maximumEntries || bytes > options.maximumBytes) {
      const oldest = entries.keys().next().value as string | undefined;
      if (oldest === undefined) return;
      const entry = entries.get(oldest);
      entries.delete(oldest);
      bytes -= entry?.bytes ?? 0;
      evictions += 1;
    }
  };

  return {
    get(key) {
      const entry = entries.get(key);
      if (entry === undefined) {
        misses += 1;
        return undefined;
      }
      entries.delete(key);
      entries.set(key, entry);
      hits += 1;
      return entry.value;
    },
    set(key, value) {
      const previous = entries.get(key);
      if (previous !== undefined) {
        entries.delete(key);
        bytes -= previous.bytes;
      }
      const size = byteSize(value);
      entries.set(key, { value, bytes: size });
      bytes += size;
      evict();
    },
    clear() {
      entries.clear();
      bytes = 0;
    },
    stats() {
      return Object.freeze({ entries: entries.size, bytes, hits, misses, evictions });
    },
  };
}
