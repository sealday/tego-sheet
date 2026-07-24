import type { JsonValue } from '../../core/types/json';

const maximumDepth = 64;

/** Error raised when a value cannot cross the strict JSON-safe SDK boundary. */
export class JsonSnapshotError extends TypeError {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'JsonSnapshotError';
  }
}

function arrayIndex(key: string, length: number): number | undefined {
  if (!/^(?:0|[1-9]\d*)$/.test(key)) return undefined;
  const index = Number(key);
  return Number.isSafeInteger(index) && index < length ? index : undefined;
}

function cloneJsonValue(
  value: unknown,
  ancestors: WeakSet<object>,
  depth: number,
  path: string,
): JsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new JsonSnapshotError(`${path} must be a finite number`);
    return value;
  }
  if (typeof value !== 'object') {
    throw new JsonSnapshotError(`${path} must contain only JSON values`);
  }
  if (depth > maximumDepth) {
    throw new JsonSnapshotError(`${path} exceeds the maximum JSON nesting depth`);
  }
  if (ancestors.has(value)) throw new JsonSnapshotError(`${path} contains a cycle`);
  ancestors.add(value);
  try {
    if (Array.isArray(value)) {
      if (Object.getPrototypeOf(value) !== Array.prototype) {
        throw new JsonSnapshotError(`${path} must use a plain array prototype`);
      }
      const keys = Reflect.ownKeys(value);
      if (
        keys.some(
          (key) =>
            typeof key !== 'string' ||
            (key !== 'length' && arrayIndex(key, value.length) === undefined),
        )
      ) {
        throw new JsonSnapshotError(`${path} has unsupported array properties`);
      }
      const result: JsonValue[] = [];
      for (let index = 0; index < value.length; index += 1) {
        const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
        if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
          throw new JsonSnapshotError(`${path}[${index}] must be a plain data property`);
        }
        result.push(cloneJsonValue(descriptor.value, ancestors, depth + 1, `${path}[${index}]`));
      }
      return Object.freeze(result);
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new JsonSnapshotError(`${path} must use a plain object prototype`);
    }
    const result: Record<string, JsonValue> = {};
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== 'string') {
        throw new JsonSnapshotError(`${path} cannot contain symbol keys`);
      }
      if (key === 'toJSON') throw new JsonSnapshotError(`${path}.toJSON is not allowed`);
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        throw new JsonSnapshotError(`${path}.${key} must be a plain data property`);
      }
      Object.defineProperty(result, key, {
        value: cloneJsonValue(descriptor.value, ancestors, depth + 1, `${path}.${key}`),
        enumerable: true,
        configurable: false,
        writable: false,
      });
    }
    return Object.freeze(result);
  } catch (cause) {
    if (cause instanceof JsonSnapshotError) throw cause;
    throw new JsonSnapshotError(`${path} could not be inspected safely`, { cause });
  } finally {
    ancestors.delete(value);
  }
}

/** Clones and deeply freezes a strict JSON value without invoking user code. */
export function snapshotJsonValue(value: unknown, path = 'value'): JsonValue {
  return cloneJsonValue(value, new WeakSet(), 0, path);
}

/** Returns the exact UTF-8 JSON encoding size of an already-safe snapshot. */
export function jsonSnapshotBytes(value: JsonValue): number {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

/** Captures a strict string array without invoking accessors or iterators. */
export function snapshotStringList(
  value: unknown,
  path: string,
  predicate: (entry: string) => boolean,
  allowEmpty: boolean,
): readonly string[] {
  const snapshot = snapshotJsonValue(value, path);
  if (
    !Array.isArray(snapshot) ||
    (!allowEmpty && snapshot.length === 0) ||
    snapshot.some((entry) => typeof entry !== 'string' || !predicate(entry)) ||
    new Set(snapshot).size !== snapshot.length
  ) {
    throw new JsonSnapshotError(`${path} must be a unique string array`);
  }
  return snapshot as readonly string[];
}
