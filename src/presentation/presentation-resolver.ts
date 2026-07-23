import type {
  Cell,
  Diagnostic,
  DocumentCellAddress,
  DocumentSheetId,
  JsonValue,
  SpreadsheetDocument,
} from '../document';
import { formulaAddressKey, type FormulaProgram, type FormulaValue } from '../formula';
import { createNumberFormatter } from '../format';
import { checkboxCellType } from '../extensions/cell-types/checkbox';
import { dropdownCellType } from '../extensions/cell-types/dropdown';
import type { BuiltInCellTypeDefinition, CellTypeScalar } from '../extensions/kernel/capabilities';
import type {
  CellPresentation,
  PresentationAnnotation,
  PresentationResolver,
  PresentationValidation,
  ResolvedStyle,
} from './cell-presentation';
import type { PresentationCache } from './presentation-cache';

const DEFAULT_STYLE: ResolvedStyle = Object.freeze({
  color: '#0a0a0a',
  backgroundColor: '#ffffff',
  fontFamily: 'Arial',
  fontSize: 10,
  bold: false,
  italic: false,
  horizontalAlign: 'left',
  verticalAlign: 'middle',
  wrap: false,
});

/** Explicit presentation target and deterministic locale inputs. */
export interface PresentationEnvironment {
  /** BCP 47 locale used for formatting. */
  readonly locale: string;
  /** IANA time-zone identifier used for deterministic formatting. */
  readonly timeZone: string;
  /** Workbook Excel serial-date system. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  /** Output channel requesting the presentation. */
  readonly target: 'screen' | 'accessibility' | 'print';
}

/** Revision tuple that invalidates presentation cache entries. */
export interface PresentationRevisions {
  /** Persistent document revision. */
  readonly document: number;
  /** Formula calculation revision. */
  readonly calculation: number;
  /** Conditional-format evaluation revision. */
  readonly condition: number;
  /** Style registry revision. */
  readonly style: number;
  /** Locale, time-zone, font, or target environment revision. */
  readonly environment: number;
}

/** Inputs required to create a deterministic presentation resolver. */
export interface PresentationResolverOptions {
  /** Immutable Workbook 2.0 source snapshot. */
  readonly document: SpreadsheetDocument;
  /** Typed F3 formula program for the same document snapshot. */
  readonly formulaProgram: FormulaProgram;
  /** Explicit deterministic presentation environment. */
  readonly environment: PresentationEnvironment;
  /** Revisions included in every cache key. */
  readonly revisions: PresentationRevisions;
  /** Explicitly budgeted presentation cache. */
  readonly cache: PresentationCache;
  /** Optional validation-state resolver. */
  readonly validation?: (address: DocumentCellAddress) => PresentationValidation;
  /** Optional ordered annotation resolver. */
  readonly annotations?: (address: DocumentCellAddress) => readonly PresentationAnnotation[];
}

function record(value: JsonValue | undefined): Readonly<Record<string, JsonValue>> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Readonly<Record<string, JsonValue>>)
    : {};
}

function string(
  value: Readonly<Record<string, JsonValue>>,
  modern: string,
  legacy: string,
): string | undefined {
  const candidate = value[modern] ?? value[legacy];
  return typeof candidate === 'string' ? candidate : undefined;
}

function boolean(value: Readonly<Record<string, JsonValue>>, key: string): boolean | undefined {
  return typeof value[key] === 'boolean' ? value[key] : undefined;
}

function number(value: Readonly<Record<string, JsonValue>>, key: string): number | undefined {
  const candidate = value[key];
  return typeof candidate === 'number' && Number.isFinite(candidate) ? candidate : undefined;
}

function style(options: PresentationResolverOptions, cell: Cell | undefined): ResolvedStyle {
  const styleEntry = options.document.workbook.styles.find(({ id }) => id === cell?.styleId);
  const value = record(styleEntry?.value);
  const font = record(value.font);
  const fontSize = number(value, 'fontSize') ?? number(font, 'size') ?? DEFAULT_STYLE.fontSize;
  const horizontalAlign = string(value, 'horizontalAlign', 'align');
  const verticalAlign = string(value, 'verticalAlign', 'valign');
  const numberFormat = string(value, 'numberFormat', 'format');
  return Object.freeze({
    color: string(value, 'color', 'color') ?? DEFAULT_STYLE.color,
    backgroundColor: string(value, 'backgroundColor', 'bgcolor') ?? DEFAULT_STYLE.backgroundColor,
    fontFamily: string(value, 'fontFamily', 'font') ?? string(font, 'name', 'name') ?? 'Arial',
    fontSize,
    bold: boolean(value, 'bold') ?? boolean(font, 'bold') ?? false,
    italic: boolean(value, 'italic') ?? boolean(font, 'italic') ?? false,
    horizontalAlign:
      horizontalAlign === 'center' || horizontalAlign === 'right' ? horizontalAlign : 'left',
    verticalAlign: verticalAlign === 'top' || verticalAlign === 'bottom' ? verticalAlign : 'middle',
    wrap: boolean(value, 'wrap') ?? boolean(value, 'textwrap') ?? false,
    ...(numberFormat === undefined ? {} : { numberFormat }),
    ...(boolean(value, 'underline') === undefined
      ? {}
      : { underline: boolean(value, 'underline') }),
    ...(boolean(value, 'strike') === undefined ? {} : { strike: boolean(value, 'strike') }),
  });
}

function inputValue(cell: Cell | undefined): FormulaValue {
  const input = cell?.input;
  if (input === undefined || input.type === 'blank') return { type: 'blank' };
  if (input.type === 'string') return { type: 'string', value: input.value };
  if (input.type === 'number') return { type: 'number', value: input.value };
  if (input.type === 'boolean') return { type: 'boolean', value: input.value };
  if (input.type === 'formula') return { type: 'blank' };
  return { type: 'error', value: '#VALUE!' };
}

interface CustomPresentation {
  readonly value: FormulaValue;
  readonly formattedText: string;
  readonly label: string;
  readonly role: 'text' | 'checkbox' | 'combobox';
  readonly checked?: boolean;
}

function scalarValue(value: CellTypeScalar): FormulaValue {
  if (value === null) return { type: 'blank' };
  if (typeof value === 'string') return { type: 'string', value };
  if (typeof value === 'number') return { type: 'number', value };
  return { type: 'boolean', value };
}

function customPresentation(
  cell: Cell | undefined,
  environment: PresentationEnvironment,
): CustomPresentation | undefined {
  if (cell?.input.type !== 'custom') return undefined;
  if (cell.input.cellType === 'checkbox')
    return describeCustom(checkboxCellType, cell.input, environment);
  if (cell.input.cellType === 'dropdown')
    return describeCustom(dropdownCellType, cell.input, environment);
  return undefined;
}

function describeCustom<Value extends JsonValue>(
  definition: BuiltInCellTypeDefinition<Value>,
  input: Extract<Cell['input'], { readonly type: 'custom' }>,
  environment: PresentationEnvironment,
): CustomPresentation | undefined {
  const value =
    input.schemaVersion === definition.schemaVersion && definition.validate(input.value)
      ? input.value
      : definition.migrate?.(input.value, input.schemaVersion);
  if (value === undefined || !definition.validate(value)) return undefined;
  const semantics = definition.describe(value, environment);
  return {
    value: scalarValue(definition.toFormulaScalar(value)),
    formattedText: semantics.formattedText,
    label: semantics.accessibilityLabel,
    role: semantics.role,
    ...(semantics.checked === undefined ? {} : { checked: semantics.checked }),
  };
}

function plain(value: FormulaValue): string {
  if (value.type === 'blank') return '';
  if (value.type === 'array') return '#SPILL!';
  return String(value.value);
}

function freezePresentation(value: CellPresentation): CellPresentation {
  return Object.freeze({
    ...value,
    address: Object.freeze({ ...value.address }),
    value: Object.freeze({ ...value.value }),
    style: Object.freeze({ ...value.style }),
    validation: Object.freeze({ ...value.validation }),
    annotations: Object.freeze(
      value.annotations.map((annotation) => Object.freeze({ ...annotation })),
    ),
    visibility: Object.freeze({ ...value.visibility }),
    accessibility: Object.freeze({ ...value.accessibility }),
    ...(value.diagnostics === undefined
      ? {}
      : {
          diagnostics: Object.freeze(
            value.diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic })),
          ),
        }),
  });
}

function addressKey(address: DocumentCellAddress, revisions: PresentationRevisions): string {
  return [
    address.sheetId,
    address.row,
    address.column,
    revisions.document,
    revisions.calculation,
    revisions.condition,
    revisions.style,
    revisions.environment,
  ].join(':');
}

function cellAt(
  document: SpreadsheetDocument,
  sheetId: DocumentSheetId,
  row: number,
  column: number,
): Cell | undefined {
  return document.workbook.sheets
    .find((sheet) => sheet.id === sheetId)
    ?.cells.find((entry) => entry.row === row && entry.column === column)?.cell;
}

/** Creates the one shared resolver used by visual, semantic and print surfaces. */
export function createPresentationResolver(
  options: PresentationResolverOptions,
): PresentationResolver {
  const formatter = createNumberFormatter();
  const resolver: PresentationResolver = {
    resolve(address: DocumentCellAddress) {
      const key = addressKey(address, options.revisions);
      const cached = options.cache.get(key);
      if (cached !== undefined) return cached;
      const sheet = options.document.workbook.sheets.find(({ id }) => id === address.sheetId);
      if (sheet === undefined) throw new RangeError(`Unknown sheet: ${address.sheetId}`);
      const cell = cellAt(options.document, address.sheetId, address.row, address.column);
      const resolvedStyle = style(options, cell);
      const formulaKey = formulaAddressKey(address);
      const formulaValue = options.formulaProgram.values.get(formulaKey);
      const custom = customPresentation(cell, options.environment);
      const value =
        custom?.value ??
        (cell?.input.type === 'formula' ? (formulaValue ?? inputValue(cell)) : inputValue(cell));
      let formattedText = custom?.formattedText ?? plain(value);
      const diagnostics: Diagnostic[] = [];
      if (custom === undefined && resolvedStyle.numberFormat !== undefined) {
        try {
          formattedText = formatter.format(value, resolvedStyle.numberFormat, options.environment);
        } catch (cause) {
          diagnostics.push({
            code: 'PRESENTATION_FORMAT_INVALID',
            severity: 'warning',
            domain: 'format',
            stage: 'resolve',
            message: 'Cell number format could not be resolved',
            location: { cell: address },
            cause,
          });
        }
      }
      const validation = options.validation?.(address) ?? { status: 'valid' };
      const annotations = options.annotations?.(address) ?? [];
      const hidden =
        sheet.rows.some((row) => row.index === address.row && row.hidden === true) ||
        sheet.columns.some((column) => column.index === address.column && column.hidden === true);
      const description = validation.status === 'valid' ? undefined : validation.message;
      const presentation = freezePresentation({
        address,
        value,
        formattedText,
        style: resolvedStyle,
        validation,
        annotations,
        visibility: { hidden, printable: cell?.printable !== false },
        accessibility: {
          label:
            description === undefined
              ? (custom?.label ?? formattedText)
              : `${custom?.label ?? formattedText}, ${description}`,
          ...(description === undefined ? {} : { description }),
          readOnly: cell?.editable === false,
          invalid: validation.status === 'error' || value.type === 'error',
          ...(custom === undefined ? {} : { role: custom.role }),
          ...(custom?.checked === undefined ? {} : { checked: custom.checked }),
        },
        ...(diagnostics.length === 0 ? {} : { diagnostics }),
      });
      options.cache.set(key, presentation);
      return presentation;
    },
  };
  return Object.freeze(resolver);
}
