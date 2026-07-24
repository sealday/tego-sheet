import type { FormulaValue } from '../formula';
import { formulaAddressKey } from '../formula';
import type { Diagnostic, Sheet, SpreadsheetDocument } from '../document';
import { resolveObjectAnchor, type ObjectGeometry } from '../objects';
import type { DisplayRect, PrintDisplayCommand } from '../print';
import {
  chartToDisplayCommands,
  resolveChart,
  type ChartDefinition,
  type ChartValueSource,
} from './charts';
import {
  resolveSparkline,
  sparklineToDisplayCommands,
  type SparklineDefinition,
} from './sparklines';
import type { AnalysisRangeReference } from './references';

export interface VisualizationPlacement {
  readonly chart: (definition: ChartDefinition) => DisplayRect | null;
  readonly sparkline: (definition: SparklineDefinition) => DisplayRect | null;
}

export interface PersistedVisualizationProjection {
  readonly id: string;
  readonly kind: 'chart' | 'sparkline';
  readonly sourceRevision: string;
  readonly rect: DisplayRect;
  readonly commands: readonly PrintDisplayCommand[];
  readonly summary: string;
  readonly diagnostics: readonly Diagnostic[];
}

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

/** Resolves persisted definitions once against an immutable revision into shared display commands. */
export function projectPersistedVisualizations(
  sheet: Pick<Sheet, 'charts' | 'sparklines'>,
  source: ChartValueSource,
  placement: VisualizationPlacement,
): readonly PersistedVisualizationProjection[] {
  const charts = sheet.charts
    .map((definition): PersistedVisualizationProjection | undefined => {
      const rect = placement.chart(definition);
      if (rect === null) return undefined;
      const resolution = resolveChart(definition, source);
      const display = chartToDisplayCommands(resolution.model, rect);
      return Object.freeze({
        id: definition.id,
        kind: 'chart',
        sourceRevision: resolution.model.sourceRevision,
        rect: Object.freeze({ ...rect }),
        commands: display.commands,
        summary: display.summary,
        diagnostics: resolution.diagnostics,
      });
    })
    .filter(
      (projection): projection is PersistedVisualizationProjection => projection !== undefined,
    );
  const sparklines = sheet.sparklines
    .map((definition): PersistedVisualizationProjection | undefined => {
      const rect = placement.sparkline(definition);
      if (rect === null) return undefined;
      const resolution = resolveSparkline(definition, source);
      const display = sparklineToDisplayCommands(resolution.model, rect);
      return Object.freeze({
        id: definition.id,
        kind: 'sparkline',
        sourceRevision: resolution.model.sourceRevision,
        rect: Object.freeze({ ...rect }),
        commands: display.commands,
        summary: display.summary,
        diagnostics: resolution.diagnostics,
      });
    })
    .filter(
      (projection): projection is PersistedVisualizationProjection => projection !== undefined,
    );
  return Object.freeze(
    [...charts, ...sparklines].sort(
      (left, right) =>
        (left.kind === right.kind ? 0 : left.kind === 'chart' ? -1 : 1) ||
        compareCodeUnits(left.id, right.id),
    ),
  );
}

/** Default worksheet placement shared by screen projections and unpaginated output. */
export function createVisualizationPlacement(geometry: ObjectGeometry): VisualizationPlacement {
  return Object.freeze({
    chart: (definition: ChartDefinition) =>
      definition.anchor === undefined ? null : resolveObjectAnchor(definition.anchor, geometry),
    sparkline: (definition: SparklineDefinition) => ({
      x: geometry.columnOffset(definition.target.column),
      y: geometry.rowOffset(definition.target.row),
      width:
        geometry.columnOffset(definition.target.column + 1) -
        geometry.columnOffset(definition.target.column),
      height:
        geometry.rowOffset(definition.target.row + 1) - geometry.rowOffset(definition.target.row),
    }),
  });
}

function scalar(value: FormulaValue | undefined): unknown {
  if (value === undefined || value.type === 'blank' || value.type === 'array') return null;
  return value.value;
}

function inputValue(
  document: SpreadsheetDocument,
  sheetId: string,
  row: number,
  column: number,
): unknown {
  const input = document.workbook.sheets
    .find((sheet) => sheet.id === sheetId)
    ?.cells.find((entry) => entry.row === row && entry.column === column)?.cell.input;
  if (input === undefined || input.type === 'blank' || input.type === 'formula') return null;
  return input.value;
}

/** Creates a bounded range source whose values and revision are captured by the caller snapshot. */
export function createPersistedVisualizationValueSource(
  document: SpreadsheetDocument,
  revision: string,
  formulaValues: ReadonlyMap<string, FormulaValue> = new Map(),
): ChartValueSource {
  return Object.freeze({
    revision,
    read(reference: AnalysisRangeReference): readonly unknown[] {
      const values: unknown[] = [];
      for (let row = reference.start.row; row <= reference.end.row; row += 1) {
        for (let column = reference.start.column; column <= reference.end.column; column += 1) {
          const calculated = formulaValues.get(
            formulaAddressKey({ sheetId: reference.sheetId, row, column }),
          );
          values.push(
            calculated === undefined
              ? inputValue(document, reference.sheetId, row, column)
              : scalar(calculated),
          );
        }
      }
      return Object.freeze(values);
    },
  });
}
