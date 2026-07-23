export type DocumentPatchPath = readonly (string | number)[];

export type DocumentPatchOperation =
  | {
      readonly op: 'set';
      readonly path: DocumentPatchPath;
      readonly value: unknown;
    }
  | {
      readonly op: 'delete';
      readonly path: DocumentPatchPath;
    }
  | {
      readonly op: 'splice';
      readonly path: DocumentPatchPath;
      readonly index: number;
      readonly deleteCount: number;
      readonly values: readonly unknown[];
    };

export interface DocumentPatch {
  readonly operations: readonly DocumentPatchOperation[];
}

function cloneValue<T>(value: T): T {
  if (Array.isArray(value)) return value.map((item) => cloneValue(item)) as T;
  if (value !== null && typeof value === 'object') {
    const output = Object.create(
      Object.getPrototypeOf(value) === null ? null : Object.prototype,
    ) as Record<string, unknown>;
    for (const key of Object.keys(value)) {
      Object.defineProperty(output, key, {
        configurable: true,
        enumerable: true,
        value: cloneValue((value as Record<string, unknown>)[key]),
        writable: true,
      });
    }
    return output as T;
  }
  return value;
}

function freezeValue<T>(value: T): T {
  if (value !== null && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const key of Object.keys(value)) {
      freezeValue((value as Record<string, unknown>)[key]);
    }
    Object.freeze(value);
  }
  return value;
}

function sameValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((item, index) => sameValue(item, right[index]));
  }
  const leftKeys = Object.keys(left).sort();
  const rightKeys = Object.keys(right).sort();
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key, index) =>
        key === rightKeys[index] &&
        sameValue((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key]),
    )
  );
}

function diffValue(
  before: unknown,
  after: unknown,
  path: readonly (string | number)[],
  operations: DocumentPatchOperation[],
): void {
  if (sameValue(before, after)) return;
  if (Array.isArray(before) && Array.isArray(after)) {
    if (before.length === after.length) {
      for (let index = 0; index < before.length; index += 1) {
        diffValue(before[index], after[index], [...path, index], operations);
      }
      return;
    }
    let prefix = 0;
    while (
      prefix < before.length &&
      prefix < after.length &&
      sameValue(before[prefix], after[prefix])
    ) {
      prefix += 1;
    }
    let suffix = 0;
    while (
      suffix < before.length - prefix &&
      suffix < after.length - prefix &&
      sameValue(before[before.length - suffix - 1], after[after.length - suffix - 1])
    ) {
      suffix += 1;
    }
    operations.push({
      op: 'splice',
      path,
      index: prefix,
      deleteCount: before.length - prefix - suffix,
      values: after.slice(prefix, after.length - suffix).map((item) => cloneValue(item)),
    });
    return;
  }
  if (
    before !== null &&
    after !== null &&
    typeof before === 'object' &&
    typeof after === 'object' &&
    !Array.isArray(before) &&
    !Array.isArray(after)
  ) {
    const beforeRecord = before as Record<string, unknown>;
    const afterRecord = after as Record<string, unknown>;
    const keys = new Set([...Object.keys(beforeRecord), ...Object.keys(afterRecord)]);
    for (const key of [...keys].sort()) {
      if (!Object.hasOwn(afterRecord, key)) {
        operations.push({ op: 'delete', path: [...path, key] });
      } else if (!Object.hasOwn(beforeRecord, key)) {
        operations.push({
          op: 'set',
          path: [...path, key],
          value: cloneValue(afterRecord[key]),
        });
      } else {
        diffValue(beforeRecord[key], afterRecord[key], [...path, key], operations);
      }
    }
    return;
  }
  operations.push({ op: 'set', path, value: cloneValue(after) });
}

/** @internal Creates an immutable forward or inverse patch without retaining either root. */
export function createDocumentPatch(before: unknown, after: unknown): DocumentPatch {
  const operations: DocumentPatchOperation[] = [];
  diffValue(before, after, [], operations);
  return freezeValue({ operations });
}

function parentAt(
  root: unknown,
  path: DocumentPatchPath,
): { parent: unknown; key: string | number } {
  if (path.length === 0) throw new TypeError('Root patch does not have a parent');
  let parent = root;
  for (const segment of path.slice(0, -1)) {
    if (parent === null || typeof parent !== 'object') {
      throw new TypeError('Patch path does not resolve to a container');
    }
    parent = (parent as Record<string | number, unknown>)[segment];
  }
  return { parent, key: path[path.length - 1] as string | number };
}

/** @internal Applies a trusted internal patch to an isolated clone of the supplied root. */
export function applyDocumentPatch<T>(input: T, patch: DocumentPatch): T {
  let output: unknown = cloneValue(input);
  for (const operation of patch.operations) {
    if (operation.path.length === 0) {
      if (operation.op === 'set') {
        output = cloneValue(operation.value);
      } else if (operation.op === 'splice' && Array.isArray(output)) {
        output.splice(
          operation.index,
          operation.deleteCount,
          ...operation.values.map((value) => cloneValue(value)),
        );
      } else {
        throw new TypeError('Patch operation cannot mutate this root');
      }
      continue;
    }
    if (operation.op === 'splice') {
      let target = output;
      for (const segment of operation.path) {
        if (target === null || typeof target !== 'object') {
          throw new TypeError('Patch splice path does not resolve to an array');
        }
        target = (target as Record<string | number, unknown>)[segment];
      }
      if (!Array.isArray(target)) throw new TypeError('Patch splice target must be an array');
      target.splice(
        operation.index,
        operation.deleteCount,
        ...operation.values.map((value) => cloneValue(value)),
      );
      continue;
    }
    const { parent, key } = parentAt(output, operation.path);
    if (parent === null || typeof parent !== 'object') {
      throw new TypeError('Patch path parent must be an object');
    }
    if (operation.op === 'delete') {
      Reflect.deleteProperty(parent, key);
    } else {
      Object.defineProperty(parent, key, {
        configurable: true,
        enumerable: true,
        value: cloneValue(operation.value),
        writable: true,
      });
    }
  }
  return output as T;
}
