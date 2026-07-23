import type { CellInput, SpreadsheetDocument } from '../document';
import type { FormulaAst, FormulaDiagnostic, FormulaValue, ScalarFormulaValue } from './ast';
import { createDependencyGraph, formulaAddressKey, transitiveDependents } from './dependency-graph';
import type { FormulaAddress, FormulaDependencyGraph } from './dependency-graph';
import type { FormulaFunctionContext, FormulaFunctionRegistry } from './function-registry';
import { createFormulaFunctionRegistry } from './function-registry';
import { parseFormula } from './parser';
import { resolveFormulaReferences } from './reference-resolver';

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
}

/** One cell whose input or structure changed. */
export type DependencyChange = FormulaAddress;

/** Rebuildable formula AST, dependency, diagnostic, and value cache. */
export interface FormulaProgram {
  /** Immutable source document. */
  readonly document: SpreadsheetDocument;
  /** Resolved formula AST snapshots keyed by stable address. */
  readonly formulas: ReadonlyMap<string, FormulaAst>;
  /** Compiled dependency edges. */
  readonly graph: FormulaDependencyGraph;
  /** Latest calculated values. */
  readonly values: ReadonlyMap<string, FormulaValue>;
  /** Latest compile and evaluation diagnostics. */
  readonly diagnostics: ReadonlyMap<string, readonly FormulaDiagnostic[]>;
}

interface FormulaProgramState {
  readonly formulas: Map<string, FormulaAst>;
  readonly graph: FormulaDependencyGraph;
  readonly values: Map<string, FormulaValue>;
  diagnostics: Map<string, readonly FormulaDiagnostic[]>;
  lastTick?: number;
  initialized: boolean;
}

const formulaProgramStates = new WeakMap<FormulaProgram, FormulaProgramState>();

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

function scalar(value: FormulaValue): ScalarFormulaValue {
  return value.type === 'array' ? { type: 'error', value: '#VALUE!' } : value;
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

  return {
    compile(document) {
      const formulas = new Map<string, FormulaAst>();
      const diagnostics = new Map<string, readonly FormulaDiagnostic[]>();
      for (const sheet of document.workbook.sheets) {
        for (const { row, column, cell } of sheet.cells) {
          if (cell.input.type !== 'formula') continue;
          const address = formulaAddressKey({ sheetId: sheet.id, row, column });
          try {
            const resolution = resolveFormulaReferences(
              parseFormula(cell.input.source),
              document,
              sheet.id,
            );
            formulas.set(address, resolution.ast);
            if (resolution.diagnostics.length > 0) diagnostics.set(address, resolution.diagnostics);
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
          values.set(address, { type: 'error', value: '#VALUE!' });
        }
      }
      const state: FormulaProgramState = {
        formulas,
        graph,
        values,
        diagnostics,
        initialized: false,
      };
      const program: FormulaProgram = Object.freeze({
        document,
        get formulas() {
          return new Map(formulas);
        },
        get graph() {
          return graph;
        },
        get values() {
          return new Map(state.values);
        },
        get diagnostics() {
          return new Map(state.diagnostics);
        },
      });
      formulaProgramStates.set(program, state);
      return program;
    },

    recalculate(program_, changes, environment) {
      const program = formulaProgramStates.get(program_);
      if (program === undefined) throw new TypeError('FormulaProgram belongs to another engine');
      const inputs = cellInputs(program_.document);
      const all = new Set(program.formulas.keys());
      let dirty: Set<string>;
      if (!program.initialized) dirty = all;
      else {
        dirty = new Set(transitiveDependents(program.graph, changes.map(formulaAddressKey)));
        if (program.lastTick !== environment.tick) {
          for (const [address, ast] of program.formulas) {
            if (containsVolatile(ast, registry)) dirty.add(address);
          }
        }
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
      const calculationNow = environment.clock.now();

      const resolveAddress = (address: string): FormulaValue => {
        if (!dirty.has(address) && values.has(address)) return values.get(address) as FormulaValue;
        if (cycleAddresses.has(address)) {
          const value: FormulaValue = { type: 'error', value: '#REF!' };
          values.set(address, value);
          return value;
        }
        const ast = program.formulas.get(address);
        if (ast === undefined) return inputValue(inputs.get(address) ?? { type: 'blank' });
        if (evaluating.has(address)) return { type: 'error', value: '#REF!' };
        evaluationCount += 1;
        if (evaluationCount > maximumEvaluations) {
          const value: FormulaValue = { type: 'error', value: '#NUM!' };
          values.set(address, value);
          program.diagnostics.set(address, [
            {
              code: 'FORMULA_EVALUATION_LIMIT_EXCEEDED',
              message: `Calculation exceeded ${maximumEvaluations} formula evaluations`,
            },
          ]);
          return value;
        }
        evaluating.add(address);
        const value = evaluateAst(ast);
        evaluating.delete(address);
        values.set(address, value);
        evaluated.push(address);
        return value;
      };

      const evaluateReference = (sheetId: string | undefined, row: number, column: number) =>
        sheetId === undefined
          ? ({ type: 'error', value: '#REF!' } as FormulaValue)
          : resolveAddress(formulaAddressKey({ sheetId, row, column }));

      const evaluateAst = (ast: FormulaAst): FormulaValue => {
        if (ast.kind === 'number') return { type: 'number', value: ast.value };
        if (ast.kind === 'string') return { type: 'string', value: ast.value };
        if (ast.kind === 'boolean') return { type: 'boolean', value: ast.value };
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
        if (definition.mode === 'async') return { type: 'error', value: '#N/A' };
        const arguments_: ScalarFormulaValue[] = [];
        for (const argument of ast.arguments) {
          const value = evaluateAst(argument);
          if (value.type === 'array') {
            for (const row of value.rows) arguments_.push(...row);
          } else {
            if (error(value)) return value;
            arguments_.push(value);
          }
        }
        if (
          arguments_.length < definition.parameters.minimum ||
          (definition.parameters.maximum !== undefined &&
            arguments_.length > definition.parameters.maximum)
        ) {
          return { type: 'error', value: '#VALUE!' };
        }
        const context: FormulaFunctionContext = Object.freeze({
          locale: environment.locale,
          timeZone: environment.timeZone,
          dateSystem: environment.dateSystem,
          now: calculationNow,
        });
        const result = definition.evaluate(Object.freeze(arguments_), context);
        return result instanceof Promise ? { type: 'error', value: '#N/A' } : result;
      };

      for (const address of [...dirty].sort()) resolveAddress(address);
      program.initialized = true;
      program.lastTick = environment.tick;
      return {
        values: new Map(values),
        evaluatedAddresses: Object.freeze([...new Set(evaluated)]),
        cycles,
        diagnostics: new Map(program.diagnostics),
      };
    },
  };
}
