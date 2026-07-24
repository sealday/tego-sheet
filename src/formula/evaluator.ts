import type { CellInput, DocumentSheetId, SpreadsheetDocument } from '../document';
import type { FormulaAst, FormulaDiagnostic, FormulaValue, ScalarFormulaValue } from './ast';
import { createDependencyGraph, formulaAddressKey, transitiveDependents } from './dependency-graph';
import type { FormulaAddress, FormulaDependencyGraph } from './dependency-graph';
import type { FormulaFunctionContext, FormulaFunctionRegistry } from './function-registry';
import { createFormulaFunctionRegistry } from './function-registry';
import { parseFormula } from './parser';
import { resolveFormulaReferences } from './reference-resolver';
import {
  planFormulaSpill,
  type FormulaNameRegistry,
  type FormulaTableBindingResolver,
} from './advanced';

/** Explicit deterministic inputs for one recalculation. */
export interface CalculationEnvironment {
  /** BCP 47 locale used by declared functions. */
  readonly locale: string;
  /** IANA time-zone identifier used by declared functions. */
  readonly timeZone: string;
  /** Workbook Excel serial-date system. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  /** Injected clock; the engine samples it once per recalculation. */
  readonly clock: {
    /** Returns the explicitly controlled Unix timestamp in milliseconds. */
    readonly now: () => number;
  };
  /** Host-controlled invalidation token for volatile functions. */
  readonly tick: number;
  /** Version expected for the supplied function registry. */
  readonly functionRegistryVersion: string;
  /** Whether volatile functions may use the explicitly supplied clock and time zone. */
  readonly resolveVolatile?: boolean;
}

/** One cell whose input or structure changed. */
export interface DependencyChange extends FormulaAddress {
  /** Optional replacement input from the committed transaction snapshot. */
  readonly input?: CellInput;
}

/** Rebuildable formula AST, dependency, diagnostic, and value cache. */
export interface FormulaProgram {
  /** Immutable source document. */
  readonly document: SpreadsheetDocument;
  /** Resolved formula AST snapshots keyed by stable address. */
  readonly formulas: ReadonlyMap<string, FormulaAst>;
  /** Compiled dependency edges. */
  readonly graph: FormulaDependencyGraph;
  /** Projected spill children mapped to their formula anchor. */
  readonly spillAnchors: ReadonlyMap<string, string>;
  /** Latest calculated values. */
  readonly values: ReadonlyMap<string, FormulaValue>;
  /** Latest compile and evaluation diagnostics. */
  readonly diagnostics: ReadonlyMap<string, readonly FormulaDiagnostic[]>;
}

interface FormulaProgramState {
  formulas: Map<string, FormulaAst>;
  baseGraph: FormulaDependencyGraph;
  graph: FormulaDependencyGraph;
  inputs: Map<string, CellInput>;
  values: Map<string, FormulaValue>;
  diagnostics: Map<string, readonly FormulaDiagnostic[]>;
  anchors: Map<string, FormulaAddress>;
  spills: Map<string, ReadonlySet<string>>;
  lastTick?: number;
  lastFunctionRegistryVersion?: string;
  initialized: boolean;
}

/** Result of one deterministic incremental recalculation. */
export interface CalculationResult {
  /** Snapshot of calculated values after the run. */
  readonly values: ReadonlyMap<string, FormulaValue>;
  /** Stable ordered formula addresses actually evaluated. */
  readonly evaluatedAddresses: readonly string[];
  /** Stable minimal strongly connected formula components. */
  readonly cycles: readonly (readonly string[])[];
  /** Compile and evaluation diagnostics keyed by address. */
  readonly diagnostics: ReadonlyMap<string, readonly FormulaDiagnostic[]>;
}

/** Compiler and incremental evaluator for Workbook 2.0 formulas. */
export interface FormulaEngine {
  /** Compiles formula source and dependency edges from a document snapshot. */
  compile(document: SpreadsheetDocument): FormulaProgram;
  /** Recalculates dirty formulas with explicit deterministic inputs. */
  recalculate(
    program: FormulaProgram,
    changes: readonly DependencyChange[],
    environment: CalculationEnvironment,
  ): CalculationResult;
}

/** Formula-engine construction options. */
export interface FormulaEngineOptions {
  /** Isolated function registry; defaults to the built-in registry. */
  readonly functions?: FormulaFunctionRegistry;
  /** Maximum formula evaluations allowed in one recalculation. */
  readonly maximumEvaluations?: number;
  /** Maximum AST and range-cell evaluation steps allowed in one recalculation. */
  readonly maximumCalculationSteps?: number;
  /** Optional stable named-range registry used during parsing and dependency binding. */
  readonly names?: FormulaNameRegistry;
  /** Injectable structured-table binding used by the evaluator before TBL-01 persistence lands. */
  readonly tables?: FormulaTableBindingResolver;
  /** Maximum projected cells for one dynamic-array formula. */
  readonly maximumSpillCells?: number;
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

function quotedSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

function bindNames(
  source: string,
  document: SpreadsheetDocument,
  currentSheetId: string,
  names: FormulaNameRegistry | undefined,
): string {
  if (names === undefined) return source;
  return source.replace(/\b[A-Za-z_][A-Za-z0-9_.]*\b/gu, (token, offset: number) => {
    const following = source.slice(offset + token.length).trimStart()[0];
    if (following === '(') return token;
    const definition = names.resolve(token, currentSheetId);
    if (definition === undefined) return token;
    const sheet = document.workbook.sheets.find(({ id }) => id === definition.refersTo.sheetId);
    if (sheet === undefined) return '#REF!';
    const start = `${columnLabel(definition.refersTo.start.column)}${definition.refersTo.start.row + 1}`;
    const end = `${columnLabel(definition.refersTo.end.column)}${definition.refersTo.end.row + 1}`;
    const reference = start === end ? start : `${start}:${end}`;
    return `${quotedSheetName(sheet.name)}!${reference}`;
  });
}

function bindStructuredReferences(
  source: string,
  document: SpreadsheetDocument,
  currentSheetId: string,
  tables: FormulaTableBindingResolver | undefined,
): { readonly source: string; readonly diagnostics: readonly FormulaDiagnostic[] } {
  if (tables === undefined) return { source, diagnostics: [] };
  const diagnostics: FormulaDiagnostic[] = [];
  const bound = source.replace(
    /\b([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]/gu,
    (match, tableName: string, columnName: string, offset: number) => {
      const result = tables.resolve({ tableName, columnName, currentSheetId });
      if (result.status === 'invalid') {
        diagnostics.push({
          code: 'FORMULA_REFERENCE_INVALID',
          message: result.message,
          span: { start: offset, end: offset + match.length },
        });
        return '#REF!';
      }
      const sheet = document.workbook.sheets.find(({ id }) => id === result.range.sheetId);
      if (sheet === undefined) {
        diagnostics.push({
          code: 'FORMULA_REFERENCE_INVALID',
          message: `Structured reference ${tableName}[${columnName}] targets an unknown sheet`,
          span: { start: offset, end: offset + match.length },
        });
        return '#REF!';
      }
      const start = `${columnLabel(result.range.start.column)}${result.range.start.row + 1}`;
      const end = `${columnLabel(result.range.end.column)}${result.range.end.row + 1}`;
      return `${quotedSheetName(sheet.name)}!${start === end ? start : `${start}:${end}`}`;
    },
  );
  return { source: bound, diagnostics };
}

function bindFormulaSource(
  source: string,
  document: SpreadsheetDocument,
  currentSheetId: string,
  options: Pick<FormulaEngineOptions, 'names' | 'tables'>,
): { readonly source: string; readonly diagnostics: readonly FormulaDiagnostic[] } {
  const structured = bindStructuredReferences(source, document, currentSheetId, options.tables);
  return {
    source: bindNames(structured.source, document, currentSheetId, options.names),
    diagnostics: structured.diagnostics,
  };
}

function inputValue(input: CellInput): ScalarFormulaValue {
  if (input.type === 'blank') return { type: 'blank' };
  if (input.type === 'number') return { type: 'number', value: input.value };
  if (input.type === 'string') return { type: 'string', value: input.value };
  if (input.type === 'boolean') return { type: 'boolean', value: input.value };
  return input.type === 'formula' ? { type: 'blank' } : { type: 'error', value: '#VALUE!' };
}

function cellInputs(document: SpreadsheetDocument): Map<string, CellInput> {
  const inputs = new Map<string, CellInput>();
  for (const sheet of document.workbook.sheets) {
    for (const { row, column, cell } of sheet.cells) {
      inputs.set(formulaAddressKey({ sheetId: sheet.id, row, column }), cell.input);
    }
  }
  return inputs;
}

function snapshotGraph(graph: FormulaDependencyGraph): FormulaDependencyGraph {
  return {
    dependencies: new Map(
      [...graph.dependencies].map(([address, dependencies]) => [address, new Set(dependencies)]),
    ),
    dependents: new Map(
      [...graph.dependents].map(([address, dependents]) => [address, new Set(dependents)]),
    ),
  };
}

function spillAnchorMap(
  spills: ReadonlyMap<string, ReadonlySet<string>>,
): ReadonlyMap<string, string> {
  const anchors = new Map<string, string>();
  for (const [anchor, projected] of spills) {
    for (const address of projected) {
      if (address !== anchor) anchors.set(address, anchor);
    }
  }
  return anchors;
}

function remapSpillDependencies(
  graph: FormulaDependencyGraph,
  spills: ReadonlyMap<string, ReadonlySet<string>>,
): FormulaDependencyGraph {
  const anchors = spillAnchorMap(spills);
  if (anchors.size === 0) return snapshotGraph(graph);
  const dependencies = new Map<string, ReadonlySet<string>>();
  const dependents = new Map<string, Set<string>>();
  for (const [formula, references] of graph.dependencies) {
    const remapped = new Set(
      [...references].map((reference) => anchors.get(reference) ?? reference),
    );
    dependencies.set(formula, remapped);
    for (const reference of remapped) {
      const formulas = dependents.get(reference) ?? new Set<string>();
      formulas.add(formula);
      dependents.set(reference, formulas);
    }
  }
  return { dependencies, dependents };
}

function snapshotDiagnostics(
  diagnostics: ReadonlyMap<string, readonly FormulaDiagnostic[]>,
): ReadonlyMap<string, readonly FormulaDiagnostic[]> {
  return new Map(
    [...diagnostics].map(([address, values]) => [
      address,
      values.map((diagnostic) => ({
        ...diagnostic,
        ...(diagnostic.span === undefined ? {} : { span: { ...diagnostic.span } }),
      })),
    ]),
  );
}

function scalar(value: FormulaValue): ScalarFormulaValue {
  return value.type === 'array' ? { type: 'error', value: '#VALUE!' } : value;
}

function snapshotFormulaValue(value: FormulaValue): FormulaValue {
  if (value.type !== 'array') return Object.freeze({ ...value });
  return Object.freeze({
    type: 'array',
    rows: Object.freeze(
      value.rows.map((row) => Object.freeze(row.map((item) => Object.freeze({ ...item })))),
    ),
  });
}

function error(value: ScalarFormulaValue): value is Extract<ScalarFormulaValue, { type: 'error' }> {
  return value.type === 'error';
}

function number(value: ScalarFormulaValue): number | undefined {
  if (value.type === 'number') return value.value;
  if (value.type === 'boolean') return value.value ? 1 : 0;
  if (value.type === 'blank') return 0;
  if (value.type === 'string') {
    const converted = Number(value.value);
    return Number.isFinite(converted) ? converted : undefined;
  }
  return undefined;
}

function compare(
  operator: Extract<FormulaAst, { kind: 'binary' }>['operator'],
  left: ScalarFormulaValue,
  right: ScalarFormulaValue,
): FormulaValue {
  const leftValue = left.type === 'blank' ? '' : left.value;
  const rightValue = right.type === 'blank' ? '' : right.value;
  if (operator === '&')
    return { type: 'string', value: `${String(leftValue)}${String(rightValue)}` };
  const leftNumber = number(left);
  const rightNumber = number(right);
  if (['=', '==', '<>', '!=', '>', '>=', '<', '<='].includes(operator)) {
    const first =
      leftNumber === undefined || rightNumber === undefined ? String(leftValue) : leftNumber;
    const second =
      leftNumber === undefined || rightNumber === undefined ? String(rightValue) : rightNumber;
    let result = false;
    if (operator === '=' || operator === '==') result = first === second;
    else if (operator === '<>' || operator === '!=') result = first !== second;
    else if (operator === '>') result = first > second;
    else if (operator === '>=') result = first >= second;
    else if (operator === '<') result = first < second;
    else result = first <= second;
    return { type: 'boolean', value: result };
  }
  if (leftNumber === undefined || rightNumber === undefined)
    return { type: 'error', value: '#VALUE!' };
  if (operator === '/' && rightNumber === 0) return { type: 'error', value: '#DIV/0!' };
  const result =
    operator === '+'
      ? leftNumber + rightNumber
      : operator === '-'
        ? leftNumber - rightNumber
        : operator === '*'
          ? leftNumber * rightNumber
          : leftNumber / rightNumber;
  return Number.isFinite(result)
    ? { type: 'number', value: result }
    : { type: 'error', value: '#NUM!' };
}

function findCycles(
  graph: FormulaDependencyGraph,
  formulas: ReadonlyMap<string, FormulaAst>,
): readonly (readonly string[])[] {
  let index = 0;
  const indexes = new Map<string, number>();
  const low = new Map<string, number>();
  const stack: string[] = [];
  const onStack = new Set<string>();
  const cycles: string[][] = [];
  const connect = (node: string): void => {
    indexes.set(node, index);
    low.set(node, index);
    index += 1;
    stack.push(node);
    onStack.add(node);
    for (const dependency of graph.dependencies.get(node) ?? []) {
      if (!formulas.has(dependency)) continue;
      if (!indexes.has(dependency)) {
        connect(dependency);
        low.set(node, Math.min(low.get(node) as number, low.get(dependency) as number));
      } else if (onStack.has(dependency)) {
        low.set(node, Math.min(low.get(node) as number, indexes.get(dependency) as number));
      }
    }
    if (low.get(node) !== indexes.get(node)) return;
    const component: string[] = [];
    let member: string;
    do {
      member = stack.pop() as string;
      onStack.delete(member);
      component.push(member);
    } while (member !== node);
    const selfLoop = component.length === 1 && graph.dependencies.get(node)?.has(node);
    if (component.length > 1 || selfLoop) cycles.push(component.sort());
  };
  for (const node of [...formulas.keys()].sort()) if (!indexes.has(node)) connect(node);
  return cycles.sort((left, right) => (left[0] as string).localeCompare(right[0] as string));
}

function containsVolatile(ast: FormulaAst, registry: FormulaFunctionRegistry): boolean {
  if (ast.kind === 'call') {
    if (registry.resolve(ast.name)?.volatility === 'volatile') return true;
    return ast.arguments.some((argument) => containsVolatile(argument, registry));
  }
  if (ast.kind === 'binary')
    return containsVolatile(ast.left, registry) || containsVolatile(ast.right, registry);
  if (ast.kind === 'unary') return containsVolatile(ast.operand, registry);
  return false;
}

/** Creates an isolated formula compiler and incremental evaluator. */
export function createFormulaEngine(options: FormulaEngineOptions = {}): FormulaEngine {
  const registry = options.functions ?? createFormulaFunctionRegistry();
  const maximumEvaluations = options.maximumEvaluations ?? 100_000;
  const maximumCalculationSteps = options.maximumCalculationSteps ?? 1_000_000;
  const maximumSpillCells = options.maximumSpillCells ?? 100_000;
  const formulaProgramStates = new WeakMap<FormulaProgram, FormulaProgramState>();

  return {
    compile(document) {
      const formulas = new Map<string, FormulaAst>();
      const anchors = new Map<string, FormulaAddress>();
      const diagnostics = new Map<string, readonly FormulaDiagnostic[]>();
      for (const sheet of document.workbook.sheets) {
        for (const { row, column, cell } of sheet.cells) {
          if (cell.input.type !== 'formula') continue;
          const address = formulaAddressKey({ sheetId: sheet.id, row, column });
          anchors.set(address, { sheetId: sheet.id, row, column });
          try {
            const binding = bindFormulaSource(cell.input.source, document, sheet.id, options);
            const resolution = resolveFormulaReferences(
              parseFormula(binding.source),
              document,
              sheet.id,
            );
            formulas.set(address, resolution.ast);
            const combined = [...binding.diagnostics, ...resolution.diagnostics];
            if (combined.length > 0) diagnostics.set(address, combined);
          } catch (cause) {
            diagnostics.set(address, [
              {
                code: 'FORMULA_PARSE_ERROR',
                message: cause instanceof Error ? cause.message : 'Formula parse failed',
              },
            ]);
          }
        }
      }
      const graph = createDependencyGraph(formulas);
      const values = new Map<string, FormulaValue>();
      for (const [address, formulaDiagnostics] of diagnostics) {
        if (formulaDiagnostics.some(({ code }) => code === 'FORMULA_PARSE_ERROR')) {
          values.set(address, snapshotFormulaValue({ type: 'error', value: '#VALUE!' }));
        }
      }
      const state: FormulaProgramState = {
        formulas,
        baseGraph: graph,
        graph,
        inputs: cellInputs(document),
        values,
        diagnostics,
        anchors,
        spills: new Map(),
        initialized: false,
      };
      const program: FormulaProgram = Object.freeze({
        document,
        get formulas() {
          return new Map(state.formulas);
        },
        get graph() {
          return snapshotGraph(state.graph);
        },
        get spillAnchors() {
          return new Map(spillAnchorMap(state.spills));
        },
        get values() {
          return new Map(state.values);
        },
        get diagnostics() {
          return snapshotDiagnostics(state.diagnostics);
        },
      });
      formulaProgramStates.set(program, state);
      return program;
    },

    recalculate(program_, changes, environment) {
      const storedProgram = formulaProgramStates.get(program_);
      if (storedProgram === undefined)
        throw new TypeError('FormulaProgram belongs to another engine');
      if (environment.functionRegistryVersion !== registry.version) {
        throw new TypeError(
          `Formula registry version ${environment.functionRegistryVersion} does not match ${registry.version}`,
        );
      }
      const program: FormulaProgramState = {
        formulas: new Map(storedProgram.formulas),
        baseGraph: storedProgram.baseGraph,
        graph: storedProgram.graph,
        inputs: new Map(storedProgram.inputs),
        values: new Map(storedProgram.values),
        diagnostics: new Map(storedProgram.diagnostics),
        anchors: new Map(storedProgram.anchors),
        spills: new Map(storedProgram.spills),
        ...(storedProgram.lastTick === undefined ? {} : { lastTick: storedProgram.lastTick }),
        ...(storedProgram.lastFunctionRegistryVersion === undefined
          ? {}
          : { lastFunctionRegistryVersion: storedProgram.lastFunctionRegistryVersion }),
        initialized: storedProgram.initialized,
      };
      const changedKeys = changes.map(formulaAddressKey);
      const previouslyAffected = transitiveDependents(program.graph, changedKeys);
      const changedFormulas = new Set<string>();
      let graphChanged = false;
      for (const change of changes) {
        if (change.input === undefined) continue;
        const address = formulaAddressKey(change);
        program.inputs.set(address, change.input);
        if (change.input.type === 'formula') {
          program.anchors.set(address, {
            sheetId: change.sheetId,
            row: change.row,
            column: change.column,
          });
          try {
            const binding = bindFormulaSource(
              change.input.source,
              program_.document,
              change.sheetId,
              options,
            );
            const resolution = resolveFormulaReferences(
              parseFormula(binding.source),
              program_.document,
              change.sheetId,
            );
            program.formulas.set(address, resolution.ast);
            const combined = [...binding.diagnostics, ...resolution.diagnostics];
            if (combined.length > 0) program.diagnostics.set(address, combined);
            else program.diagnostics.delete(address);
          } catch (cause) {
            program.formulas.delete(address);
            program.values.set(address, snapshotFormulaValue({ type: 'error', value: '#VALUE!' }));
            program.diagnostics.set(address, [
              {
                code: 'FORMULA_PARSE_ERROR',
                message: cause instanceof Error ? cause.message : 'Formula parse failed',
              },
            ]);
          }
          changedFormulas.add(address);
          graphChanged = true;
        } else if (program.formulas.delete(address)) {
          program.anchors.delete(address);
          program.values.delete(address);
          graphChanged = true;
        }
      }
      if (graphChanged) {
        program.baseGraph = createDependencyGraph(program.formulas);
        program.graph = program.baseGraph;
      }

      let dirty: Set<string>;
      if (!program.initialized) dirty = new Set(program.formulas.keys());
      else {
        dirty = new Set(previouslyAffected);
        if (program.lastFunctionRegistryVersion !== environment.functionRegistryVersion) {
          for (const address of program.formulas.keys()) dirty.add(address);
        }
        if (graphChanged) {
          for (const address of transitiveDependents(program.graph, changedKeys))
            dirty.add(address);
        }
        for (const address of changedFormulas) dirty.add(address);
        if (program.lastTick !== environment.tick) {
          for (const [address, ast] of program.formulas) {
            if (containsVolatile(ast, registry)) dirty.add(address);
          }
        }
      }
      const changedKeySet = new Set(changedKeys);
      for (const [anchor, projected] of program.spills) {
        if (!dirty.has(anchor) && ![...projected].some((address) => changedKeySet.has(address))) {
          continue;
        }
        for (const address of projected) program.values.delete(address);
        program.spills.delete(anchor);
        dirty.add(anchor);
      }
      for (const [address, value] of program.values) {
        if (value.type === 'error' && value.value === '#SPILL!') dirty.add(address);
      }
      for (const [address, diagnostics] of program.diagnostics) {
        const retained = diagnostics.filter(
          ({ code }) =>
            code !== 'FORMULA_CIRCULAR_REFERENCE' &&
            (!dirty.has(address) ||
              ![
                'FORMULA_UNKNOWN_FUNCTION',
                'FORMULA_EVALUATION_LIMIT_EXCEEDED',
                'VOLATILE_FORMULA_NOT_RESOLVED',
                'ASYNC_FORMULA_NOT_RESOLVED',
              ].includes(code)),
        );
        if (retained.length === 0) program.diagnostics.delete(address);
        else if (retained.length !== diagnostics.length) program.diagnostics.set(address, retained);
      }
      const cycles = findCycles(program.graph, program.formulas);
      const cycleAddresses = new Set(cycles.flat());
      for (const cycle of cycles) {
        for (const address of cycle) {
          program.diagnostics.set(address, [
            {
              code: 'FORMULA_CIRCULAR_REFERENCE',
              message: `Circular reference: ${cycle.join(' -> ')}`,
            },
          ]);
        }
      }
      const values = program.values;
      const evaluated: string[] = [];
      const evaluating = new Set<string>();
      let evaluationCount = 0;
      let calculationSteps = 0;
      const calculationNow = environment.clock.now();

      const resolveAddress = (address: string): FormulaValue => {
        if (!dirty.has(address) && values.has(address)) return values.get(address) as FormulaValue;
        if (cycleAddresses.has(address)) {
          const value: FormulaValue = { type: 'error', value: '#REF!' };
          const snapshot = snapshotFormulaValue(value);
          values.set(address, snapshot);
          return snapshot;
        }
        const ast = program.formulas.get(address);
        if (ast === undefined) {
          const input = program.inputs.get(address);
          if (input?.type === 'formula' && values.has(address))
            return values.get(address) as FormulaValue;
          return inputValue(input ?? { type: 'blank' });
        }
        if (evaluating.has(address)) return { type: 'error', value: '#REF!' };
        evaluationCount += 1;
        if (evaluationCount > maximumEvaluations) {
          const value: FormulaValue = { type: 'error', value: '#NUM!' };
          const snapshot = snapshotFormulaValue(value);
          values.set(address, snapshot);
          program.diagnostics.set(address, [
            {
              code: 'FORMULA_EVALUATION_LIMIT_EXCEEDED',
              message: `Calculation exceeded ${maximumEvaluations} formula evaluations`,
            },
          ]);
          return snapshot;
        }
        evaluating.add(address);
        const value = evaluateAst(ast);
        evaluating.delete(address);
        let snapshot = snapshotFormulaValue(value);
        if (snapshot.type === 'array') {
          const anchor = program.anchors.get(address);
          if (anchor === undefined) {
            snapshot = { type: 'error', value: '#SPILL!' };
          } else {
            const occupied = new Set(
              [...program.inputs]
                .filter(
                  ([key, input]) =>
                    key !== address &&
                    input.type !== 'blank' &&
                    !(program.spills.get(address)?.has(key) ?? false),
                )
                .map(([key]) => key),
            );
            for (const [otherAnchor, projected] of program.spills) {
              if (otherAnchor === address) continue;
              for (const key of projected) occupied.add(key);
            }
            const plan = planFormulaSpill({
              anchor: { ...anchor, sheetId: anchor.sheetId as DocumentSheetId },
              value: snapshot,
              occupied,
              limits: { maxCells: maximumSpillCells },
            });
            if (plan.status === 'blocked') {
              snapshot = plan.value;
            } else {
              const projected = new Set<string>();
              for (const [key, projectedValue] of plan.cells) {
                values.set(key, snapshotFormulaValue(projectedValue));
                projected.add(key);
              }
              program.spills.set(address, projected);
              snapshot = values.get(address) ?? { type: 'blank' };
            }
          }
        }
        values.set(address, snapshot);
        evaluated.push(address);
        return snapshot;
      };

      const evaluateReference = (sheetId: string | undefined, row: number, column: number) =>
        sheetId === undefined
          ? ({ type: 'error', value: '#REF!' } as FormulaValue)
          : resolveAddress(formulaAddressKey({ sheetId, row, column }));

      const evaluateAst = (ast: FormulaAst): FormulaValue => {
        calculationSteps += 1;
        if (calculationSteps > maximumCalculationSteps) {
          const address = [...evaluating].at(-1);
          if (address !== undefined) {
            program.diagnostics.set(address, [
              {
                code: 'FORMULA_EVALUATION_LIMIT_EXCEEDED',
                message: `Calculation exceeded ${maximumCalculationSteps} evaluation steps`,
                span: ast.span,
              },
            ]);
          }
          return { type: 'error', value: '#NUM!' };
        }
        if (ast.kind === 'number') return { type: 'number', value: ast.value };
        if (ast.kind === 'string') return { type: 'string', value: ast.value };
        if (ast.kind === 'boolean') return { type: 'boolean', value: ast.value };
        if (ast.kind === 'error') return { type: 'error', value: ast.value };
        if (ast.kind === 'reference') {
          return evaluateReference(ast.reference.sheetId, ast.reference.row, ast.reference.column);
        }
        if (ast.kind === 'range') {
          if (ast.start.sheetId === undefined || ast.start.sheetId !== ast.end.sheetId)
            return { type: 'error', value: '#REF!' };
          const rows: ScalarFormulaValue[][] = [];
          for (
            let row = Math.min(ast.start.row, ast.end.row);
            row <= Math.max(ast.start.row, ast.end.row);
            row += 1
          ) {
            const values_: ScalarFormulaValue[] = [];
            for (
              let column = Math.min(ast.start.column, ast.end.column);
              column <= Math.max(ast.start.column, ast.end.column);
              column += 1
            ) {
              calculationSteps += 1;
              if (calculationSteps > maximumCalculationSteps) {
                const address = [...evaluating].at(-1);
                if (address !== undefined) {
                  program.diagnostics.set(address, [
                    {
                      code: 'FORMULA_EVALUATION_LIMIT_EXCEEDED',
                      message: `Calculation exceeded ${maximumCalculationSteps} evaluation steps`,
                      span: ast.span,
                    },
                  ]);
                }
                return { type: 'error', value: '#NUM!' };
              }
              values_.push(scalar(evaluateReference(ast.start.sheetId, row, column)));
            }
            rows.push(values_);
          }
          return { type: 'array', rows };
        }
        if (ast.kind === 'unary') {
          const operand = scalar(evaluateAst(ast.operand));
          if (error(operand)) return operand;
          const converted = number(operand);
          return converted === undefined
            ? { type: 'error', value: '#VALUE!' }
            : { type: 'number', value: -converted };
        }
        if (ast.kind === 'binary') {
          const left = scalar(evaluateAst(ast.left));
          if (error(left)) return left;
          const right = scalar(evaluateAst(ast.right));
          if (error(right)) return right;
          return compare(ast.operator, left, right);
        }
        const definition = registry.resolve(ast.name);
        if (definition === undefined) {
          const address = [...evaluating].at(-1);
          if (address !== undefined) {
            program.diagnostics.set(address, [
              {
                code: 'FORMULA_UNKNOWN_FUNCTION',
                message: `Unknown formula function ${ast.name}`,
                span: ast.span,
              },
            ]);
          }
          return { type: 'error', value: '#NAME?' };
        }
        if (
          ast.arguments.length < definition.parameters.minimum ||
          (definition.parameters.maximum !== undefined &&
            ast.arguments.length > definition.parameters.maximum)
        ) {
          return { type: 'error', value: '#VALUE!' };
        }
        if (definition.volatility === 'volatile' && environment.resolveVolatile === false) {
          const address = [...evaluating].at(-1);
          if (address !== undefined) {
            program.diagnostics.set(address, [
              {
                code: 'VOLATILE_FORMULA_NOT_RESOLVED',
                message: `Volatile formula function ${ast.name} requires an explicit clock and time zone`,
                span: ast.span,
              },
            ]);
          }
          return { type: 'error', value: '#N/A' };
        }
        if (definition.mode === 'async') {
          const address = [...evaluating].at(-1);
          if (address !== undefined) {
            program.diagnostics.set(address, [
              {
                code: 'ASYNC_FORMULA_NOT_RESOLVED',
                message: `Async formula function ${ast.name} has no fixed result`,
                span: ast.span,
              },
            ]);
          }
          return { type: 'error', value: '#N/A' };
        }
        if (ast.name === 'IF') {
          const condition = scalar(
            evaluateAst(ast.arguments[0] ?? { kind: 'boolean', value: false, span: ast.span }),
          );
          if (error(condition)) return condition;
          const selected = (
            condition.type === 'blank'
              ? false
              : condition.type === 'boolean'
                ? condition.value
                : number(condition)
          )
            ? ast.arguments[1]
            : ast.arguments[2];
          return selected === undefined ? { type: 'blank' } : evaluateAst(selected);
        }
        const arguments_: ScalarFormulaValue[] = [];
        for (const argument of ast.arguments) {
          const value = evaluateAst(argument);
          if (value.type === 'array') {
            for (const row of value.rows) {
              for (const item of row) {
                if (error(item)) return item;
                arguments_.push(item);
              }
            }
          } else {
            if (error(value)) return value;
            arguments_.push(value);
          }
        }
        const context: FormulaFunctionContext = Object.freeze({
          locale: environment.locale,
          timeZone: environment.timeZone,
          dateSystem: environment.dateSystem,
          now: calculationNow,
        });
        const frozenArguments = Object.freeze(
          arguments_.map((value) => Object.freeze({ ...value })),
        );
        const result = definition.evaluate(frozenArguments, context);
        return result instanceof Promise ? { type: 'error', value: '#N/A' } : result;
      };

      for (const address of [...dirty].sort()) resolveAddress(address);
      program.graph = remapSpillDependencies(program.baseGraph, program.spills);
      program.initialized = true;
      program.lastTick = environment.tick;
      program.lastFunctionRegistryVersion = environment.functionRegistryVersion;
      storedProgram.formulas = program.formulas;
      storedProgram.baseGraph = program.baseGraph;
      storedProgram.graph = program.graph;
      storedProgram.inputs = program.inputs;
      storedProgram.values = program.values;
      storedProgram.diagnostics = program.diagnostics;
      storedProgram.anchors = program.anchors;
      storedProgram.spills = program.spills;
      storedProgram.initialized = program.initialized;
      storedProgram.lastTick = program.lastTick;
      storedProgram.lastFunctionRegistryVersion = program.lastFunctionRegistryVersion;
      return {
        values: new Map(values),
        evaluatedAddresses: Object.freeze([...new Set(evaluated)]),
        cycles,
        diagnostics: snapshotDiagnostics(program.diagnostics),
      };
    },
  };
}
