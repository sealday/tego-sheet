import type { JsonValue } from '../../core/types/json';

const maximumDepth = 64;
const maximumNodes = 100_000;
const maximumProperties = 100_000;
const maximumStringBytes = 8 * 1024 * 1024;
const maximumEstimatedBytes = 16 * 1024 * 1024;

export interface JsonSnapshotLimits {
  readonly maxNodes: number;
  readonly maxProperties: number;
  readonly maxStringBytes: number;
  readonly maxEstimatedBytes: number;
}

type JsonSnapshotLimitResource = 'nodes' | 'properties' | 'stringBytes' | 'estimatedBytes';

interface JsonSnapshotBudget {
  readonly path: string;
  readonly limits: JsonSnapshotLimits;
  nodes: number;
  properties: number;
  stringBytes: number;
  estimatedBytes: number;
}

/** Error raised when a value cannot cross the strict JSON-safe SDK boundary. */
export class JsonSnapshotError extends TypeError {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'JsonSnapshotError';
  }
}

/** Error raised before an untrusted JSON snapshot can exceed a cumulative resource budget. */
export class JsonSnapshotLimitError extends JsonSnapshotError {
  readonly resource: JsonSnapshotLimitResource;
  readonly actual: number;
  readonly limit: number;

  constructor(path: string, resource: JsonSnapshotLimitResource, actual: number, limit: number) {
    super(`${path} exceeds the maximum JSON snapshot ${resource}`);
    this.name = 'JsonSnapshotLimitError';
    this.resource = resource;
    this.actual = actual;
    this.limit = limit;
  }
}

function snapshotLimits(overrides: Partial<JsonSnapshotLimits> | undefined): JsonSnapshotLimits {
  const limits = Object.freeze({
    maxNodes: overrides?.maxNodes ?? maximumNodes,
    maxProperties: overrides?.maxProperties ?? maximumProperties,
    maxStringBytes: overrides?.maxStringBytes ?? maximumStringBytes,
    maxEstimatedBytes: overrides?.maxEstimatedBytes ?? maximumEstimatedBytes,
  });
  if (
    !Number.isSafeInteger(limits.maxNodes) ||
    limits.maxNodes <= 0 ||
    !Number.isSafeInteger(limits.maxProperties) ||
    limits.maxProperties <= 0 ||
    !Number.isSafeInteger(limits.maxStringBytes) ||
    limits.maxStringBytes <= 0 ||
    !Number.isSafeInteger(limits.maxEstimatedBytes) ||
    limits.maxEstimatedBytes <= 0
  ) {
    throw new JsonSnapshotError('JSON snapshot limits must be positive safe integers');
  }
  return limits;
}

function assertLimit(
  budget: JsonSnapshotBudget,
  resource: JsonSnapshotLimitResource,
  actual: number,
): void {
  const limit =
    resource === 'nodes'
      ? budget.limits.maxNodes
      : resource === 'properties'
        ? budget.limits.maxProperties
        : resource === 'stringBytes'
          ? budget.limits.maxStringBytes
          : budget.limits.maxEstimatedBytes;
  if (actual > limit) {
    throw new JsonSnapshotLimitError(budget.path, resource, actual, limit);
  }
}

function addToBudget(
  budget: JsonSnapshotBudget,
  resource: JsonSnapshotLimitResource,
  amount: number,
): void {
  const actual = budget[resource] + amount;
  assertLimit(budget, resource, actual);
  budget[resource] = actual;
}

function utf8StringBytes(value: string): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x7f) bytes += 1;
    else if (code <= 0x7ff) bytes += 2;
    else if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        bytes += 4;
        index += 1;
      } else {
        bytes += 3;
      }
    } else {
      bytes += 3;
    }
  }
  return bytes;
}

function jsonStringBytes(value: string): number {
  let bytes = 2;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x22 || code === 0x5c || code === 0x08 || code === 0x09) {
      bytes += 2;
    } else if (code === 0x0a || code === 0x0c || code === 0x0d) {
      bytes += 2;
    } else if (code <= 0x1f) {
      bytes += 6;
    } else if (code <= 0x7f) {
      bytes += 1;
    } else if (code <= 0x7ff) {
      bytes += 2;
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        bytes += 4;
        index += 1;
      } else {
        bytes += 6;
      }
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      bytes += 6;
    } else {
      bytes += 3;
    }
  }
  return bytes;
}

function arrayIndex(key: string, length: number): number | undefined {
  if (!/^(?:0|[1-9]\d*)$/.test(key)) return undefined;
  const index = Number(key);
  return Number.isSafeInteger(index) && index < length ? index : undefined;
}

function cloneJsonValue(
  value: unknown,
  ancestors: WeakSet<object>,
  budget: JsonSnapshotBudget,
  depth: number,
  path: string,
): JsonValue {
  addToBudget(budget, 'nodes', 1);
  if (value === null) {
    addToBudget(budget, 'estimatedBytes', 4);
    return value;
  }
  if (typeof value === 'string') {
    addToBudget(budget, 'stringBytes', utf8StringBytes(value));
    addToBudget(budget, 'estimatedBytes', jsonStringBytes(value));
    return value;
  }
  if (typeof value === 'boolean') {
    addToBudget(budget, 'estimatedBytes', value ? 4 : 5);
    return value;
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new JsonSnapshotError(`${path} must be a finite number`);
    addToBudget(budget, 'estimatedBytes', JSON.stringify(value).length);
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
      assertLimit(budget, 'nodes', budget.nodes + value.length);
      addToBudget(budget, 'properties', value.length);
      const structuralBytes = value.length === 0 ? 2 : value.length + 1;
      assertLimit(budget, 'estimatedBytes', budget.estimatedBytes + structuralBytes + value.length);
      addToBudget(budget, 'estimatedBytes', structuralBytes);
      const result: JsonValue[] = [];
      for (let index = 0; index < value.length; index += 1) {
        const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
        if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
          throw new JsonSnapshotError(`${path}[${index}] must be a plain data property`);
        }
        result.push(
          cloneJsonValue(descriptor.value, ancestors, budget, depth + 1, `${path}[${index}]`),
        );
      }
      return Object.freeze(result);
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new JsonSnapshotError(`${path} must use a plain object prototype`);
    }
    const keys = Reflect.ownKeys(value);
    let structuralBytes = keys.length === 0 ? 2 : keys.length * 2 + 1;
    let keyStringBytes = 0;
    for (const key of keys) {
      if (typeof key !== 'string') {
        throw new JsonSnapshotError(`${path} cannot contain symbol keys`);
      }
      if (key === 'toJSON') throw new JsonSnapshotError(`${path}.toJSON is not allowed`);
      keyStringBytes += utf8StringBytes(key);
      structuralBytes += jsonStringBytes(key);
    }
    assertLimit(budget, 'nodes', budget.nodes + keys.length);
    addToBudget(budget, 'properties', keys.length);
    addToBudget(budget, 'stringBytes', keyStringBytes);
    assertLimit(budget, 'estimatedBytes', budget.estimatedBytes + structuralBytes + keys.length);
    addToBudget(budget, 'estimatedBytes', structuralBytes);
    const result: Record<string, JsonValue> = {};
    for (const key of keys as string[]) {
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        throw new JsonSnapshotError(`${path}.${key} must be a plain data property`);
      }
      Object.defineProperty(result, key, {
        value: cloneJsonValue(descriptor.value, ancestors, budget, depth + 1, `${path}.${key}`),
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
export function snapshotJsonValue(
  value: unknown,
  path = 'value',
  limits?: Partial<JsonSnapshotLimits>,
): JsonValue {
  const budget: JsonSnapshotBudget = {
    path,
    limits: snapshotLimits(limits),
    nodes: 0,
    properties: 0,
    stringBytes: 0,
    estimatedBytes: 0,
  };
  return cloneJsonValue(value, new WeakSet(), budget, 0, path);
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
