import type { FormulaAst, FormulaReference } from './ast';

/** Stable sheet-qualified zero-based formula address. */
export interface FormulaAddress {
  /** Stable sheet identity. */
  readonly sheetId: string;
  /** Zero-based row. */
  readonly row: number;
  /** Zero-based column. */
  readonly column: number;
}

/** Forward and reverse edges for compiled formula dependencies. */
export interface FormulaDependencyGraph {
  /** Referenced addresses by formula address. */
  readonly dependencies: ReadonlyMap<string, ReadonlySet<string>>;
  /** Formula addresses affected by each referenced address. */
  readonly dependents: ReadonlyMap<string, ReadonlySet<string>>;
}

function columnLabel(column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

/** Converts a stable address into a deterministic sheet-id/A1 key. */
export function formulaAddressKey(address: FormulaAddress): string {
  return `${address.sheetId}!${columnLabel(address.column)}${address.row + 1}`;
}

function keyOfReference(reference: FormulaReference): string | undefined {
  return reference.sheetId === undefined
    ? undefined
    : formulaAddressKey({
        sheetId: reference.sheetId,
        row: reference.row,
        column: reference.column,
      });
}

/** Collects every direct cell dependency from a resolved AST. */
export function collectFormulaDependencies(ast: FormulaAst): ReadonlySet<string> {
  const result = new Set<string>();
  const add = (address: string): void => {
    if (result.has(address)) return;
    if (result.size >= 100_000) {
      throw new RangeError('Formula dependency limit exceeds 100000 cells');
    }
    result.add(address);
  };
  const visit = (node: FormulaAst): void => {
    if (node.kind === 'reference') {
      const key = keyOfReference(node.reference);
      if (key !== undefined) add(key);
      return;
    }
    if (node.kind === 'range') {
      if (node.start.sheetId === undefined || node.end.sheetId !== node.start.sheetId) return;
      const startRow = Math.min(node.start.row, node.end.row);
      const endRow = Math.max(node.start.row, node.end.row);
      const startColumn = Math.min(node.start.column, node.end.column);
      const endColumn = Math.max(node.start.column, node.end.column);
      const size = (endRow - startRow + 1) * (endColumn - startColumn + 1);
      if (!Number.isSafeInteger(size) || size > 100_000) {
        throw new RangeError('Formula dependency limit exceeds 100000 cells');
      }
      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = startColumn; column <= endColumn; column += 1) {
          add(formulaAddressKey({ sheetId: node.start.sheetId, row, column }));
        }
      }
      return;
    }
    if (node.kind === 'unary') visit(node.operand);
    else if (node.kind === 'binary') {
      visit(node.left);
      visit(node.right);
    } else if (node.kind === 'call') {
      for (const argument of node.arguments) visit(argument);
    }
  };
  visit(ast);
  return result;
}

/** Builds deterministic forward and reverse dependency edges. */
export function createDependencyGraph(
  formulas: ReadonlyMap<string, FormulaAst>,
): FormulaDependencyGraph {
  const dependencies = new Map<string, ReadonlySet<string>>();
  const mutableDependents = new Map<string, Set<string>>();
  let edgeCount = 0;
  for (const [address, ast] of formulas) {
    const formulaDependencies = collectFormulaDependencies(ast);
    edgeCount += formulaDependencies.size;
    if (!Number.isSafeInteger(edgeCount) || edgeCount > 1_000_000) {
      throw new RangeError('Formula program dependency limit exceeds 1000000 edges');
    }
    dependencies.set(address, formulaDependencies);
    for (const dependency of formulaDependencies) {
      const dependents = mutableDependents.get(dependency) ?? new Set<string>();
      dependents.add(address);
      mutableDependents.set(dependency, dependents);
    }
  }
  return {
    dependencies,
    dependents: new Map(
      Array.from(mutableDependents).map(([address, dependents]) => [
        address,
        new Set(Array.from(dependents).sort()),
      ]),
    ),
  };
}

/** Finds all formula nodes transitively affected by changed addresses. */
export function transitiveDependents(
  graph: FormulaDependencyGraph,
  changes: readonly string[],
): ReadonlySet<string> {
  const dirty = new Set<string>();
  const queue = [...changes].sort();
  while (queue.length > 0) {
    const current = queue.shift() as string;
    for (const dependent of graph.dependents.get(current) ?? []) {
      if (dirty.has(dependent)) continue;
      dirty.add(dependent);
      queue.push(dependent);
      queue.sort();
    }
  }
  return dirty;
}
