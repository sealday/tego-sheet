import type { FormulaValue, ScalarFormulaValue } from './ast';

/** Restricted immutable context supplied to formula functions. */
export interface FormulaFunctionContext {
  /** Explicit calculation locale. */
  readonly locale: string;
  /** Explicit calculation time zone. */
  readonly timeZone: string;
  /** Explicit workbook date system. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  /** Clock sample captured once for the recalculation. */
  readonly now: number;
}

/** Declared formula-function contract accepted by the registry and F5 kernel. */
export interface FormulaFunctionDefinition {
  /** Case-insensitive function name. */
  readonly name: string;
  /** Supported inclusive argument-count bounds. */
  readonly parameters: {
    /** Minimum accepted argument count. */
    readonly minimum: number;
    /** Optional maximum accepted argument count. */
    readonly maximum?: number;
  };
  /** Declared result category. */
  readonly returns: FormulaValue['type'];
  /** Whether an explicit tick invalidates the function. */
  readonly volatility: 'stable' | 'volatile';
  /** Execution mode; unresolved asynchronous functions are rejected by synchronous output. */
  readonly mode: 'sync' | 'async';
  /** Restricted evaluator receiving frozen inputs and context. */
  readonly evaluate: (
    arguments_: readonly ScalarFormulaValue[],
    context: FormulaFunctionContext,
  ) => FormulaValue | Promise<FormulaValue>;
}

/** Isolated registry of declared formula functions. */
export interface FormulaFunctionRegistry {
  /** Version token updated after registrations change. */
  readonly version: string;
  /** Registers a unique function and returns an idempotent removal callback. */
  register(definition: FormulaFunctionDefinition): () => void;
  /** Resolves a function by case-insensitive name. */
  resolve(name: string): FormulaFunctionDefinition | undefined;
  /** Lists definitions in stable ASCII name order. */
  list(): readonly FormulaFunctionDefinition[];
}

/** One explicitly supported built-in Excel-compatible function. */
export interface FormulaFunctionCompatibility {
  /** Function name. */
  readonly name: string;
  /** Declared compatibility baseline. */
  readonly compatibility: 'excel';
  /** Supported execution mode. */
  readonly mode: 'sync';
}

function numeric(value: ScalarFormulaValue): number | undefined {
  if (value.type === 'number') return value.value;
  if (value.type === 'boolean') return value.value ? 1 : 0;
  if (value.type === 'blank') return 0;
  if (value.type === 'string') {
    const number = Number(value.value);
    return Number.isFinite(number) ? number : undefined;
  }
  return undefined;
}

function excelSerial(timestamp: number, dateSystem: 'excel-1900' | 'excel-1904'): number {
  const utcDay = Math.floor(timestamp / 86_400_000);
  const epoch = dateSystem === 'excel-1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  let serial = utcDay - Math.floor(epoch / 86_400_000);
  if (dateSystem === 'excel-1900' && serial >= 60) serial += 1;
  return serial;
}

const numberResult = (value: number): FormulaValue =>
  Number.isFinite(value)
    ? { type: 'number', value }
    : { type: 'error', value: Number.isNaN(value) ? '#VALUE!' : '#NUM!' };

const builtins: readonly FormulaFunctionDefinition[] = [
  {
    name: 'AND',
    parameters: { minimum: 1 },
    returns: 'boolean',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({ type: 'boolean', value: values.every(truthy) }),
  },
  {
    name: 'AVERAGE',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => {
      const numbers = values.map(numeric);
      return numbers.includes(undefined)
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(
            (numbers as number[]).reduce((sum, value) => sum + value, 0) / values.length,
          );
    },
  },
  {
    name: 'CONCAT',
    parameters: { minimum: 1 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({
      type: 'string',
      value: values.map((value) => (value.type === 'blank' ? '' : String(value.value))).join(''),
    }),
  },
  {
    name: 'IF',
    parameters: { minimum: 2, maximum: 3 },
    returns: 'blank',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([condition, truthyValue, falsyValue]) =>
      truthy(condition ?? { type: 'blank' })
        ? (truthyValue ?? { type: 'blank' })
        : (falsyValue ?? { type: 'blank' }),
  },
  {
    name: 'MAX',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => numberResult(Math.max(...values.map((value) => numeric(value) ?? 0))),
  },
  {
    name: 'MIN',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => numberResult(Math.min(...values.map((value) => numeric(value) ?? 0))),
  },
  {
    name: 'NOW',
    parameters: { minimum: 0, maximum: 0 },
    returns: 'number',
    volatility: 'volatile',
    mode: 'sync',
    evaluate: (_, context) => ({
      type: 'number',
      value:
        excelSerial(context.now, context.dateSystem) +
        (((context.now % 86_400_000) + 86_400_000) % 86_400_000) / 86_400_000,
    }),
  },
  {
    name: 'OR',
    parameters: { minimum: 1 },
    returns: 'boolean',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({ type: 'boolean', value: values.some(truthy) }),
  },
  {
    name: 'SUM',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) =>
      numberResult(values.reduce((sum, value) => sum + (numeric(value) ?? 0), 0)),
  },
  {
    name: 'TODAY',
    parameters: { minimum: 0, maximum: 0 },
    returns: 'number',
    volatility: 'volatile',
    mode: 'sync',
    evaluate: (_, context) => ({
      type: 'number',
      value: excelSerial(context.now, context.dateSystem),
    }),
  },
];

function truthy(value: ScalarFormulaValue): boolean {
  if (value.type === 'blank' || value.type === 'error') return false;
  return Boolean(value.value);
}

/** Stable compatibility manifest for built-in formula functions. */
export const BUILTIN_FORMULA_COMPATIBILITY: readonly FormulaFunctionCompatibility[] = Object.freeze(
  builtins.map(({ name }) => ({ name, compatibility: 'excel' as const, mode: 'sync' as const })),
);

/** Creates an isolated registry populated with the supported built-in functions. */
export function createFormulaFunctionRegistry(): FormulaFunctionRegistry {
  const definitions = new Map(builtins.map((definition) => [definition.name, definition]));
  let generation = 0;
  return {
    get version() {
      return `builtin-1${generation === 0 ? '' : `+${generation}`}`;
    },
    register(definition) {
      const name = definition.name.toUpperCase();
      if (!/^[A-Z_][A-Z0-9_.]*$/u.test(name)) throw new TypeError('Invalid function name');
      if (definitions.has(name))
        throw new TypeError(`Formula function ${name} is already registered`);
      const snapshot = Object.freeze({
        ...definition,
        name,
        parameters: Object.freeze({ ...definition.parameters }),
      });
      definitions.set(name, snapshot);
      generation += 1;
      return () => {
        if (definitions.delete(name)) generation += 1;
      };
    },
    resolve(name) {
      return definitions.get(name.toUpperCase());
    },
    list() {
      return Object.freeze(
        [...definitions.values()].sort((left, right) =>
          left.name < right.name ? -1 : left.name > right.name ? 1 : 0,
        ),
      );
    },
  };
}
