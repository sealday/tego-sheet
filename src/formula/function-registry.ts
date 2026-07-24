import type { FormulaValue, ScalarFormulaValue } from './ast';
import type { AdapterRegistryKernel } from '../extensions/kernel/registry';
import type { KernelEnvironment } from '../extensions/kernel/manifest';

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

function excelSerial(
  timestamp: number,
  dateSystem: 'excel-1900' | 'excel-1904',
  timeZone: string,
): number {
  const parts = new Map(
    new Intl.DateTimeFormat('en-US-u-nu-latn', {
      timeZone,
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hourCycle: 'h23',
    })
      .formatToParts(timestamp)
      .map(({ type, value }) => [type, Number(value)]),
  );
  const year = parts.get('year') as number;
  const month = parts.get('month') as number;
  const day = parts.get('day') as number;
  const hour = parts.get('hour') as number;
  const minute = parts.get('minute') as number;
  const second = parts.get('second') as number;
  const utcDay = Math.floor(Date.UTC(year, month - 1, day) / 86_400_000);
  const epoch = dateSystem === 'excel-1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  let serial = utcDay - Math.floor(epoch / 86_400_000);
  if (dateSystem === 'excel-1900' && serial >= 60) serial += 1;
  return serial + (hour * 3600 + minute * 60 + second) / 86_400;
}

function dateSerial(
  year: number,
  month: number,
  day: number,
  dateSystem: 'excel-1900' | 'excel-1904',
): number {
  if (![year, month, day].every(Number.isFinite)) return Number.NaN;
  const truncatedYear = Math.trunc(year);
  const excelYear =
    truncatedYear >= 0 && truncatedYear <= 1899 ? truncatedYear + 1900 : truncatedYear;
  const timestamp = Date.UTC(excelYear, Math.trunc(month) - 1, Math.trunc(day));
  const epoch = dateSystem === 'excel-1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  let serial = Math.floor((timestamp - epoch) / 86_400_000);
  if (dateSystem === 'excel-1900' && serial >= 60) serial += 1;
  return serial;
}

function serialDateParts(
  serial: number,
  dateSystem: 'excel-1900' | 'excel-1904',
): readonly [year: number, month: number, day: number] | undefined {
  if (!Number.isFinite(serial)) return undefined;
  const day = Math.floor(serial);
  if (dateSystem === 'excel-1900' && day === 60) return [1900, 2, 29];
  const adjusted = dateSystem === 'excel-1900' && day > 60 ? day - 1 : day;
  const epoch = dateSystem === 'excel-1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  const date = new Date(epoch + adjusted * 86_400_000);
  return [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()];
}

const numberResult = (value: number): FormulaValue =>
  Number.isFinite(value)
    ? { type: 'number', value }
    : { type: 'error', value: Number.isNaN(value) ? '#VALUE!' : '#NUM!' };

function numericArgument(value: ScalarFormulaValue | undefined): number | undefined {
  return value === undefined ? undefined : numeric(value);
}

function textArgument(value: ScalarFormulaValue | undefined): string {
  return value === undefined || value.type === 'blank' ? '' : String(value.value);
}

function excelRound(value: number, digits: number): number {
  const scale = 10 ** Math.trunc(digits);
  if (!Number.isFinite(scale) || scale === 0) return value;
  const scaled = value * scale;
  return (Math.sign(scaled) * Math.round(Math.abs(scaled) + Number.EPSILON)) / scale;
}

function directionalRound(value: number, digits: number, direction: 'up' | 'down'): number {
  const scale = 10 ** Math.trunc(digits);
  if (!Number.isFinite(scale) || scale === 0) return value;
  const magnitude = Math.abs(value * scale);
  const rounded = direction === 'up' ? Math.ceil(magnitude) : Math.floor(magnitude);
  return (Math.sign(value) * rounded) / scale;
}

const builtins: readonly FormulaFunctionDefinition[] = [
  {
    name: 'ABS',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => {
      const converted = numericArgument(value);
      return converted === undefined
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(Math.abs(converted));
    },
  },
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
    name: 'COUNT',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({
      type: 'number',
      value: values.filter(({ type }) => type === 'number').length,
    }),
  },
  {
    name: 'COUNTA',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({
      type: 'number',
      value: values.filter(({ type }) => type !== 'blank').length,
    }),
  },
  {
    name: 'COUNTBLANK',
    parameters: { minimum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: (values) => ({
      type: 'number',
      value: values.filter(({ type }) => type === 'blank').length,
    }),
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
    name: 'DATE',
    parameters: { minimum: 3, maximum: 3 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([year, month, day], context) => {
      const parts = [numericArgument(year), numericArgument(month), numericArgument(day)];
      return parts.includes(undefined)
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(dateSerial(parts[0]!, parts[1]!, parts[2]!, context.dateSystem));
    },
  },
  {
    name: 'DAY',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value], context) => {
      const converted = numericArgument(value);
      const parts =
        converted === undefined ? undefined : serialDateParts(converted, context.dateSystem);
      return parts === undefined
        ? { type: 'error', value: '#VALUE!' }
        : { type: 'number', value: parts[2] };
    },
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
    name: 'LEFT',
    parameters: { minimum: 1, maximum: 2 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, count]) => {
      const length = count === undefined ? 1 : numericArgument(count);
      return length === undefined || length < 0
        ? { type: 'error', value: '#VALUE!' }
        : { type: 'string', value: textArgument(value).slice(0, Math.trunc(length)) };
    },
  },
  {
    name: 'LEN',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => ({ type: 'number', value: textArgument(value).length }),
  },
  {
    name: 'LOWER',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => ({ type: 'string', value: textArgument(value).toLowerCase() }),
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
    name: 'MID',
    parameters: { minimum: 3, maximum: 3 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, start, count]) => {
      const startIndex = numericArgument(start);
      const length = numericArgument(count);
      return startIndex === undefined || length === undefined || startIndex < 1 || length < 0
        ? { type: 'error', value: '#VALUE!' }
        : {
            type: 'string',
            value: textArgument(value).slice(
              Math.trunc(startIndex) - 1,
              Math.trunc(startIndex) - 1 + Math.trunc(length),
            ),
          };
    },
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
    name: 'MONTH',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value], context) => {
      const converted = numericArgument(value);
      const parts =
        converted === undefined ? undefined : serialDateParts(converted, context.dateSystem);
      return parts === undefined
        ? { type: 'error', value: '#VALUE!' }
        : { type: 'number', value: parts[1] };
    },
  },
  {
    name: 'NOT',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'boolean',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => ({ type: 'boolean', value: !truthy(value ?? { type: 'blank' }) }),
  },
  {
    name: 'NOW',
    parameters: { minimum: 0, maximum: 0 },
    returns: 'number',
    volatility: 'volatile',
    mode: 'sync',
    evaluate: (_, context) => ({
      type: 'number',
      value: excelSerial(context.now, context.dateSystem, context.timeZone),
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
    name: 'RIGHT',
    parameters: { minimum: 1, maximum: 2 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, count]) => {
      const length = count === undefined ? 1 : numericArgument(count);
      return length === undefined || length < 0
        ? { type: 'error', value: '#VALUE!' }
        : {
            type: 'string',
            value: Math.trunc(length) === 0 ? '' : textArgument(value).slice(-Math.trunc(length)),
          };
    },
  },
  {
    name: 'ROUND',
    parameters: { minimum: 2, maximum: 2 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, digits]) => {
      const converted = numericArgument(value);
      const precision = numericArgument(digits);
      return converted === undefined || precision === undefined
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(excelRound(converted, precision));
    },
  },
  {
    name: 'ROUNDDOWN',
    parameters: { minimum: 2, maximum: 2 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, digits]) => {
      const converted = numericArgument(value);
      const precision = numericArgument(digits);
      return converted === undefined || precision === undefined
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(directionalRound(converted, precision, 'down'));
    },
  },
  {
    name: 'ROUNDUP',
    parameters: { minimum: 2, maximum: 2 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value, digits]) => {
      const converted = numericArgument(value);
      const precision = numericArgument(digits);
      return converted === undefined || precision === undefined
        ? { type: 'error', value: '#VALUE!' }
        : numberResult(directionalRound(converted, precision, 'up'));
    },
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
    name: 'TRIM',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => ({
      type: 'string',
      value: textArgument(value).trim().replace(/\s+/gu, ' '),
    }),
  },
  {
    name: 'TODAY',
    parameters: { minimum: 0, maximum: 0 },
    returns: 'number',
    volatility: 'volatile',
    mode: 'sync',
    evaluate: (_, context) => ({
      type: 'number',
      value: Math.floor(excelSerial(context.now, context.dateSystem, context.timeZone)),
    }),
  },
  {
    name: 'UPPER',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'string',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value]) => ({ type: 'string', value: textArgument(value).toUpperCase() }),
  },
  {
    name: 'YEAR',
    parameters: { minimum: 1, maximum: 1 },
    returns: 'number',
    volatility: 'stable',
    mode: 'sync',
    evaluate: ([value], context) => {
      const converted = numericArgument(value);
      const parts =
        converted === undefined ? undefined : serialDateParts(converted, context.dateSystem);
      return parts === undefined
        ? { type: 'error', value: '#VALUE!' }
        : { type: 'number', value: parts[0] };
    },
  },
];

function snapshotDefinition(
  definition: FormulaFunctionDefinition,
): Readonly<FormulaFunctionDefinition> {
  return Object.freeze({
    ...definition,
    parameters: Object.freeze({ ...definition.parameters }),
  });
}

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
  const definitions = new Map(
    builtins.map((definition) => [definition.name, snapshotDefinition(definition)]),
  );
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
      const snapshot = snapshotDefinition({ ...definition, name });
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

/**
 * Copies registered F5 formula capabilities into an isolated calculation registry.
 *
 * The returned callback removes only the copied definitions. Kernel lifecycle remains owned by
 * the caller.
 */
export function registerKernelFormulaFunctions(
  registry: FormulaFunctionRegistry,
  kernel: AdapterRegistryKernel,
  environment: KernelEnvironment,
): () => void {
  const removals: Array<() => void> = [];
  try {
    for (const { id } of kernel.list('formula-function')) {
      removals.push(registry.register(kernel.resolve('formula-function', { id, environment })));
    }
  } catch (error) {
    for (const remove of removals.reverse()) remove();
    throw error;
  }
  let active = true;
  return () => {
    if (!active) return;
    active = false;
    for (const remove of removals.reverse()) remove();
  };
}
