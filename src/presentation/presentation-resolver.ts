import type {
  Cell,
  Diagnostic,
  DocumentCellAddress,
  DocumentSheetId,
  FilterView,
  JsonValue,
  SpreadsheetDocument,
} from '../document';
import { formulaAddressKey, type FormulaProgram, type FormulaValue } from '../formula';
import {
  createConditionalFormatEvaluator,
  createNumberFormatter,
  type ConditionalRule,
} from '../format';
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
import { applyDocumentFilterView } from '../views';

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
  /** Optional session filter-view revision. */
  readonly view?: number;
}

/** Inputs required to create a deterministic presentation resolver. */
export interface PresentationResolverOptions {
  /** Immutable Workbook 2.0 source snapshot. */
  readonly document: SpreadsheetDocument;
  /** Typed F3 formula program for the same document snapshot. */
  readonly formulaProgram?: FormulaProgram;
  /** Typed F3 values supplied by a read-only controller snapshot adapter. */
  readonly formulaValues?: ReadonlyMap<string, FormulaValue>;
  /** Explicit deterministic presentation environment. */
  readonly environment: PresentationEnvironment;
  /** Revisions included in every cache key. */
  readonly revisions: PresentationRevisions;
  /** Explicitly budgeted presentation cache. */
  readonly cache: PresentationCache;
  /** Optional document or session view composed as derived row visibility. */
  readonly activeFilterView?: FilterView;
  /** Optional per-sheet document or session views for multi-sheet output. */
  readonly activeFilterViews?: readonly FilterView[];
  /** Optional validation-state resolver. */
  readonly validation?: (address: DocumentCellAddress) => PresentationValidation;
  /** Optional derived conditional-style resolver shared by every presentation target. */
  readonly conditionalStyle?: (
    address: DocumentCellAddress,
    value: FormulaValue,
    formattedText: string,
  ) => Readonly<Partial<ResolvedStyle>>;
  /** Optional resource limits for persistent worksheet conditional rules. */
  readonly conditionalFormattingLimits?: {
    /** Maximum rules evaluated for one cell. */
    readonly maxRules: number;
    /** Maximum aggregate target cells across those rules. */
    readonly maxCells: number;
  };
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

function borders(value: Readonly<Record<string, JsonValue>>): ResolvedStyle['border'] | undefined {
  const source = record(value.border);
  const output: Partial<
    Record<'top' | 'right' | 'bottom' | 'left', readonly [style: string, color?: string]>
  > = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const line = source[side];
    if (!Array.isArray(line) || typeof line[0] !== 'string') continue;
    output[side] = typeof line[1] === 'string' ? [line[0], line[1]] : [line[0]];
  }
  return Object.keys(output).length === 0 ? undefined : output;
}

function mergedStyleValue(
  options: PresentationResolverOptions,
  address: DocumentCellAddress,
  cell: Cell | undefined,
): Readonly<Record<string, JsonValue>> {
  const sheet = options.document.workbook.sheets.find(({ id }) => id === address.sheetId);
  const ids = [
    sheet?.columns.find(({ index }) => index === address.column)?.styleId,
    sheet?.rows.find(({ index }) => index === address.row)?.styleId,
    cell?.styleId,
  ];
  let output: Record<string, JsonValue> = {};
  for (const id of ids) {
    const next = record(options.document.workbook.styles.find((entry) => entry.id === id)?.value);
    output = {
      ...output,
      ...next,
      font: { ...record(output.font), ...record(next.font) },
      border: { ...record(output.border), ...record(next.border) },
    };
  }
  return output;
}

function style(
  options: PresentationResolverOptions,
  address: DocumentCellAddress,
  cell: Cell | undefined,
): ResolvedStyle {
  const value = mergedStyleValue(options, address, cell);
  const font = record(value.font);
  const fontSize = number(value, 'fontSize') ?? number(font, 'size') ?? DEFAULT_STYLE.fontSize;
  const horizontalAlign = string(value, 'horizontalAlign', 'align');
  const verticalAlign = string(value, 'verticalAlign', 'valign');
  const numberFormat = string(value, 'numberFormat', 'format');
  const border = borders(value);
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
    ...(border === undefined ? {} : { border }),
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

function addressKey(
  address: DocumentCellAddress,
  revisions: PresentationRevisions,
  environment: PresentationEnvironment,
  activeViewId?: string,
): string {
  return JSON.stringify([
    address.sheetId,
    address.row,
    address.column,
    revisions.document,
    revisions.calculation,
    revisions.condition,
    revisions.style,
    revisions.environment,
    revisions.view ?? 0,
    activeViewId ?? '',
    environment.locale,
    environment.timeZone,
    environment.dateSystem,
    environment.target,
  ]);
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

function conditionalScalar(source: string): string | number | undefined {
  const trimmed = source.trim();
  const numberValue = Number(trimmed);
  if (trimmed.length > 0 && Number.isFinite(numberValue)) return numberValue;
  const quoted = /^(["'])(.*)\1$/u.exec(trimmed);
  return quoted?.[2];
}

function documentConditionalRules(
  document: SpreadsheetDocument,
  address: DocumentCellAddress,
): readonly ConditionalRule[] {
  const sheet = document.workbook.sheets.find(({ id }) => id === address.sheetId);
  if (sheet === undefined) return [];
  return (sheet.conditionalFormatting ?? []).flatMap((rule, index): readonly ConditionalRule[] => {
    if (rule.type === 'color-scale') {
      return [
        {
          id: `sheet-rule-${index}`,
          priority: index,
          stopIfTrue: false,
          ranges: [rule.range],
          condition: { type: 'not-blank' },
          effect: {
            type: 'color-scale',
            minimumColor: rule.minimumColor,
            ...(rule.midpointColor === undefined ? {} : { midpointColor: rule.midpointColor }),
            maximumColor: rule.maximumColor,
          },
        },
      ];
    }
    const value = conditionalScalar(rule.formula);
    const value2 = rule.formula2 === undefined ? undefined : conditionalScalar(rule.formula2);
    const patch: Record<string, JsonValue> = {};
    if (rule.style.color !== undefined) patch.color = rule.style.color;
    if (rule.style.backgroundColor !== undefined) {
      patch.backgroundColor = rule.style.backgroundColor;
    }
    if (rule.style.bold !== undefined) patch.bold = rule.style.bold;
    return [
      {
        id: `sheet-rule-${index}`,
        priority: index,
        stopIfTrue: false,
        ranges: [rule.range],
        condition:
          value === undefined
            ? {
                type: 'cell-is-formula',
                operator: rule.operator,
                source: rule.formula,
                ...(rule.formula2 === undefined ? {} : { source2: rule.formula2 }),
              }
            : {
                type: 'cell-is',
                operator: rule.operator,
                value,
                ...(value2 === undefined ? {} : { value2 }),
              },
        effect: { type: 'style', patch },
      },
    ];
  });
}

/** Creates the one shared resolver used by visual, semantic and print surfaces. */
export function createPresentationResolver(
  options: PresentationResolverOptions,
): PresentationResolver {
  const formatter = createNumberFormatter();
  const conditionalFormatter = createConditionalFormatEvaluator(
    options.conditionalFormattingLimits ?? { maxRules: 10_000, maxCells: 10_000_000 },
  );
  const conditionalRulesBySheet = new Map(
    options.document.workbook.sheets.map((sheet) => [
      sheet.id,
      documentConditionalRules(options.document, {
        sheetId: sheet.id,
        row: 0,
        column: 0,
      }),
    ]),
  );
  const conditionalLookup = (target: DocumentCellAddress): FormulaValue | undefined => {
    const targetCell = cellAt(options.document, target.sheetId, target.row, target.column);
    const calculated =
      options.formulaValues?.get(formulaAddressKey(target)) ??
      options.formulaProgram?.values.get(formulaAddressKey(target));
    return targetCell?.input.type === 'formula' ? calculated : inputValue(targetCell);
  };
  const resolveConditionalSheetId = (sheetToken: string): DocumentSheetId | undefined => {
    const exact = options.document.workbook.sheets.find(({ name }) => name === sheetToken);
    if (exact !== undefined) return exact.id;
    const insensitive = options.document.workbook.sheets.filter(
      ({ name }) => name.toLowerCase() === sheetToken.toLowerCase(),
    );
    return insensitive.length === 1 ? insensitive[0]?.id : undefined;
  };
  const activeViews =
    options.activeFilterViews ??
    [options.activeFilterView].filter((view): view is FilterView => view !== undefined);
  const filterViewHiddenRows = new Map(
    activeViews.map((view) => [
      view.range.sheetId,
      applyDocumentFilterView({
        document: options.document,
        formulaValues: options.formulaValues ?? options.formulaProgram?.values,
        view,
        locale: options.environment.locale,
        limits: { maxRows: Math.max(1, view.range.end.row - view.range.start.row + 1) },
      }).hiddenRows,
    ]),
  );
  const activeViewKey = activeViews.map(({ id }) => id).join(',');
  const resolver: PresentationResolver = {
    resolve(address: DocumentCellAddress) {
      const key = addressKey(address, options.revisions, options.environment, activeViewKey);
      const cached = options.cache.get(key);
      if (cached !== undefined) return cached;
      const sheet = options.document.workbook.sheets.find(({ id }) => id === address.sheetId);
      if (sheet === undefined) throw new RangeError(`Unknown sheet: ${address.sheetId}`);
      const cell = cellAt(options.document, address.sheetId, address.row, address.column);
      const baseStyle = style(options, address, cell);
      const formulaKey = formulaAddressKey(address);
      const formulaValue =
        options.formulaValues?.get(formulaKey) ?? options.formulaProgram?.values.get(formulaKey);
      const custom = customPresentation(cell, options.environment);
      const value =
        custom?.value ??
        (cell?.input.type === 'formula' ? (formulaValue ?? inputValue(cell)) : inputValue(cell));
      let formattedText = custom?.formattedText ?? plain(value);
      const diagnostics: Diagnostic[] = [];
      if (custom === undefined && baseStyle.numberFormat !== undefined) {
        try {
          formattedText = formatter.format(value, baseStyle.numberFormat, options.environment);
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
      const conditional = conditionalFormatter.evaluate({
        address,
        value,
        text: formattedText,
        baseStyle,
        rules: conditionalRulesBySheet.get(address.sheetId) ?? [],
        lookup: conditionalLookup,
        resolveSheetId: resolveConditionalSheetId,
      });
      diagnostics.push(
        ...conditional.diagnostics.map(
          (diagnostic): Diagnostic => ({
            code: diagnostic.code,
            severity: 'warning',
            domain: 'format',
            stage: 'resolve',
            message: `Conditional format ${diagnostic.ruleId} is not supported by presentation`,
            location: { cell: address },
          }),
        ),
      );
      const resolvedStyle = Object.freeze({
        ...baseStyle,
        ...conditional.stylePatch,
        ...options.conditionalStyle?.(address, value, formattedText),
      });
      const validation = options.validation?.(address) ?? { status: 'valid' };
      const annotations = options.annotations?.(address) ?? [];
      const hidden =
        sheet.rows.some((row) => row.index === address.row && row.hidden === true) ||
        sheet.columns.some((column) => column.index === address.column && column.hidden === true) ||
        (filterViewHiddenRows.get(address.sheetId)?.has(address.row) ?? false);
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
