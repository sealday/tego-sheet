import type { Diagnostic, ObjectAnchor } from '../../document';
import type { DisplayRect, PrintDisplayCommand } from '../../print';
import {
  checkedIdentifier,
  checkedPositiveLimit,
  rangeCellCount,
  rangesIntersect,
  snapshotRange,
  type AnalysisRangeReference,
} from '../references';

/** Chart families implemented by the renderer-neutral first-party layout. */
export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'combo';

/** One persistent chart series backed by a stable worksheet range. */
export interface ChartSeriesDefinition {
  readonly id: string;
  readonly name?: string;
  readonly values: AnalysisRangeReference;
}

/** Persistent renderer-neutral chart definition. */
export interface ChartDefinition {
  readonly id: string;
  readonly type: ChartType;
  readonly title?: string;
  readonly categories?: AnalysisRangeReference;
  readonly series: readonly ChartSeriesDefinition[];
  /** Optional persistent worksheet-object anchor used by screen and output projections. */
  readonly anchor?: ObjectAnchor;
  /** Explicit behavior when the chart intersects a repeated template region. */
  readonly templateRepeat?: 'shared' | 'per-item' | 'forbidden';
}

/** Bounded scalar range reader evaluated against one immutable revision. */
export interface ChartValueSource {
  readonly revision: string;
  readonly read: (reference: AnalysisRangeReference) => readonly unknown[];
}

/** One resolved, immutable numeric chart series. */
export interface NormalizedChartSeries {
  readonly id: string;
  readonly name: string;
  readonly values: readonly (number | null)[];
}

/** Immutable chart model shared by screen, print, and output renderers. */
export interface NormalizedChart {
  readonly id: string;
  readonly type: ChartType;
  readonly title?: string;
  readonly sourceRevision: string;
  readonly categories: readonly string[];
  readonly series: readonly NormalizedChartSeries[];
}

/** Hard limits for one chart resolution. */
export interface ChartResolutionLimits {
  readonly maximumSeries: number;
  readonly maximumPoints: number;
}

/** Stable result of resolving chart data. */
export interface ChartResolution {
  readonly model: NormalizedChart;
  readonly dependencies: readonly AnalysisRangeReference[];
  readonly diagnostics: readonly Diagnostic[];
}

/** Vector display-list commands plus screen-reader-equivalent content. */
export interface ChartDisplayOutput {
  readonly commands: readonly PrintDisplayCommand[];
  readonly summary: string;
}

const defaults: Readonly<ChartResolutionLimits> = Object.freeze({
  maximumSeries: 100,
  maximumPoints: 10_000,
});
const palette = Object.freeze(['#2563eb', '#dc2626', '#16a34a', '#9333ea', '#ea580c', '#0891b2']);

function diagnostic(code: string, message: string, details?: Diagnostic['details']): Diagnostic {
  return Object.freeze({
    code,
    severity: 'warning',
    domain: 'analysis',
    stage: 'resolve',
    message,
    ...(details === undefined ? {} : { details }),
  });
}

function chartType(value: ChartType): ChartType {
  if (!['column', 'bar', 'line', 'area', 'pie', 'scatter', 'combo'].includes(value)) {
    throw new TypeError('Chart type is unsupported');
  }
  return value;
}

function readBounded(
  source: ChartValueSource,
  reference: AnalysisRangeReference,
  expected: number,
): readonly unknown[] {
  const values = source.read(reference);
  if (!Array.isArray(values)) throw new TypeError('Chart source must return an array');
  return values.slice(0, expected);
}

/** Resolves stable range references into a bounded immutable chart snapshot. */
export function resolveChart(
  definition: ChartDefinition,
  source: ChartValueSource,
  limits?: Partial<ChartResolutionLimits>,
): ChartResolution {
  const maximumSeries = checkedPositiveLimit(
    limits?.maximumSeries,
    defaults.maximumSeries,
    'Chart series limit',
  );
  const maximumPoints = checkedPositiveLimit(
    limits?.maximumPoints,
    defaults.maximumPoints,
    'Chart point limit',
  );
  const id = checkedIdentifier(definition.id, 'Chart ID');
  const type = chartType(definition.type);
  const dependencies: AnalysisRangeReference[] = [];
  const diagnostics: Diagnostic[] = [];
  const categoryReference =
    definition.categories === undefined
      ? undefined
      : snapshotRange(definition.categories, 'Chart category range');
  if (categoryReference !== undefined) dependencies.push(categoryReference);
  const accepted = definition.series.slice(0, maximumSeries);
  if (definition.series.length > maximumSeries) {
    diagnostics.push(
      diagnostic(
        'CHART_SERIES_LIMIT_EXCEEDED',
        `Chart ${id} contains ${definition.series.length} series; only ${maximumSeries} were resolved`,
        { chartId: id, actual: definition.series.length, maximum: maximumSeries },
      ),
    );
  }
  const seriesIds = new Set<string>();
  const seriesReferences = accepted.map((series, index) => {
    const seriesId = checkedIdentifier(series.id, `Chart series ${index} ID`);
    if (seriesIds.has(seriesId)) {
      throw new TypeError(`Duplicate chart series ID ${seriesId}`);
    }
    seriesIds.add(seriesId);
    const reference = snapshotRange(series.values, `Chart series ${series.id} range`);
    dependencies.push(reference);
    return reference;
  });
  let categories: readonly string[] = [];
  if (categoryReference !== undefined) {
    const count = rangeCellCount(categoryReference);
    if (count > maximumPoints) {
      diagnostics.push(
        diagnostic(
          'CHART_POINT_LIMIT_EXCEEDED',
          `Chart ${id} category range exceeds the ${maximumPoints} point limit`,
          { chartId: id, actual: count, maximum: maximumPoints },
        ),
      );
    } else {
      categories = Object.freeze(
        readBounded(source, categoryReference, count).map((value) =>
          value === null || value === undefined ? '' : String(value),
        ),
      );
    }
  }
  let remainingPoints = maximumPoints;
  const normalizedSeries = accepted.map((series, seriesIndex): NormalizedChartSeries => {
    const reference = seriesReferences[seriesIndex]!;
    const count = rangeCellCount(reference);
    if (count > remainingPoints) {
      diagnostics.push(
        diagnostic(
          'CHART_POINT_LIMIT_EXCEEDED',
          `Chart series ${series.id} cannot fit within the ${maximumPoints} point budget`,
          {
            chartId: id,
            seriesId: series.id,
            actual: count,
            remaining: remainingPoints,
            maximum: maximumPoints,
          },
        ),
      );
      return Object.freeze({
        id: series.id,
        name: series.name ?? series.id,
        values: Object.freeze([]),
      });
    }
    remainingPoints -= count;
    const values = readBounded(source, reference, count).map((value, pointIndex) => {
      if (typeof value === 'number' && Number.isFinite(value)) return value;
      if (value === null || value === undefined || value === '') return null;
      diagnostics.push(
        diagnostic(
          'CHART_NON_NUMERIC_VALUE',
          `Chart series ${series.id} point ${pointIndex} is not numeric`,
          { chartId: id, seriesId: series.id, pointIndex },
        ),
      );
      return null;
    });
    return Object.freeze({
      id: series.id,
      name: series.name ?? series.id,
      values: Object.freeze(values),
    });
  });
  const model = Object.freeze({
    id,
    type,
    ...(definition.title === undefined ? {} : { title: definition.title }),
    sourceRevision: source.revision,
    categories,
    series: Object.freeze(normalizedSeries),
  });
  return Object.freeze({
    model,
    dependencies: Object.freeze(dependencies),
    diagnostics: Object.freeze(diagnostics),
  });
}

/** Returns whether any changed range intersects a chart dependency. */
export function chartAffectedByChanges(
  definition: ChartDefinition,
  changes: readonly AnalysisRangeReference[],
): boolean {
  const dependencies = [
    ...(definition.categories === undefined ? [] : [definition.categories]),
    ...definition.series.map(({ values }) => values),
  ];
  return dependencies.some((dependency) =>
    changes.some((change) => rangesIntersect(dependency, change)),
  );
}

function finiteRect(rect: DisplayRect): void {
  if (
    !Number.isFinite(rect.x) ||
    !Number.isFinite(rect.y) ||
    !Number.isFinite(rect.width) ||
    !Number.isFinite(rect.height) ||
    rect.width < 0 ||
    rect.height < 0
  ) {
    throw new RangeError('Chart rectangle must be finite and non-negative');
  }
}

function bounds(series: readonly NormalizedChartSeries[]): { minimum: number; maximum: number } {
  const values = series.flatMap(({ values }) =>
    values.filter((value): value is number => value !== null),
  );
  if (values.length === 0) return { minimum: 0, maximum: 1 };
  return {
    minimum: Math.min(0, ...values),
    maximum: Math.max(0, ...values),
  };
}

function chartSummary(model: NormalizedChart): string {
  const title = model.title ?? `Chart ${model.id}`;
  const series = model.series
    .map(({ name, values }) => {
      const points = values
        .map(
          (value, index) =>
            `${model.categories[index] ?? `Point ${index + 1}`}: ${value ?? 'blank'}`,
        )
        .join(', ');
      return `${name}: ${points}`;
    })
    .join('. ');
  return `${title}. ${model.type} chart. ${series}`.trim();
}

function plotRect(rect: DisplayRect, titled: boolean): DisplayRect {
  const top = rect.y + (titled ? 28 : 8);
  return {
    x: rect.x + 36,
    y: top,
    width: Math.max(0, rect.width - 48),
    height: Math.max(0, rect.y + rect.height - top - 22),
  };
}

function titleCommand(model: NormalizedChart, rect: DisplayRect): PrintDisplayCommand[] {
  if (model.title === undefined) return [];
  return [
    {
      kind: 'text',
      text: model.title,
      x: rect.x + rect.width / 2,
      y: rect.y + 18,
      maxWidth: Math.max(0, rect.width - 16),
      fontFamily: 'sans-serif',
      fontSize: 12,
      color: '#111827',
      horizontalAlign: 'center',
    },
  ];
}

function cartesianCommands(
  model: NormalizedChart,
  plot: DisplayRect,
): readonly PrintDisplayCommand[] {
  const commands: PrintDisplayCommand[] = [
    {
      kind: 'line',
      x1: plot.x,
      y1: plot.y + plot.height,
      x2: plot.x + plot.width,
      y2: plot.y + plot.height,
      color: '#6b7280',
      width: 1,
    },
    {
      kind: 'line',
      x1: plot.x,
      y1: plot.y,
      x2: plot.x,
      y2: plot.y + plot.height,
      color: '#6b7280',
      width: 1,
    },
  ];
  const pointCount = Math.max(1, ...model.series.map(({ values }) => values.length));
  const { minimum, maximum } = bounds(model.series);
  const span = maximum === minimum ? 1 : maximum - minimum;
  const mapY = (value: number): number =>
    plot.y + plot.height - ((value - minimum) / span) * plot.height;
  const mapX = (value: number): number => plot.x + ((value - minimum) / span) * plot.width;
  if (model.type === 'line' || model.type === 'area' || model.type === 'scatter') {
    model.series.forEach((series, seriesIndex) => {
      const points: string[] = [];
      const areaPoints: { x: number; y: number }[] = [];
      let startsSegment = true;
      series.values.forEach((value, pointIndex) => {
        if (value === null) {
          startsSegment = true;
          return;
        }
        const x = plot.x + ((pointIndex + 0.5) / pointCount) * plot.width;
        const y = mapY(value);
        areaPoints.push({ x, y });
        points.push(`${startsSegment ? 'M' : 'L'} ${x} ${y}`);
        startsSegment = false;
        commands.push({
          kind: 'fill-rect',
          rect: { x: x - 1.5, y: y - 1.5, width: 3, height: 3 },
          color: palette[seriesIndex % palette.length]!,
        });
      });
      if (model.type === 'area' && areaPoints.length > 0) {
        const first = areaPoints[0]!;
        const last = areaPoints[areaPoints.length - 1]!;
        commands.push({
          kind: 'path',
          data: `M ${first.x} ${mapY(0)} ${areaPoints
            .map(({ x, y }) => `L ${x} ${y}`)
            .join(' ')} L ${last.x} ${mapY(0)} Z`,
          fill: palette[seriesIndex % palette.length],
        });
      } else if (model.type !== 'scatter' && points.length > 0) {
        commands.push({
          kind: 'path',
          data: points.join(' '),
          stroke: palette[seriesIndex % palette.length],
          width: 2,
        });
      }
    });
    return commands;
  }
  const seriesCount = Math.max(1, model.series.length);
  model.series.forEach((series, seriesIndex) => {
    series.values.forEach((value, pointIndex) => {
      if (value === null) return;
      if (model.type === 'combo' && seriesIndex % 2 === 1) {
        const x = plot.x + ((pointIndex + 0.5) / pointCount) * plot.width;
        const y = mapY(value);
        commands.push({
          kind: 'fill-rect',
          rect: { x: x - 1.5, y: y - 1.5, width: 3, height: 3 },
          color: palette[seriesIndex % palette.length]!,
        });
        if (pointIndex > 0) {
          const previous = series.values[pointIndex - 1];
          if (previous !== null && previous !== undefined) {
            commands.push({
              kind: 'line',
              x1: plot.x + ((pointIndex - 0.5) / pointCount) * plot.width,
              y1: mapY(previous),
              x2: x,
              y2: y,
              color: palette[seriesIndex % palette.length]!,
              width: 2,
            });
          }
        }
        return;
      }
      if (model.type === 'bar') {
        const band = plot.height / pointCount;
        const thickness = (band * 0.72) / seriesCount;
        const zero = mapX(0);
        const edge = mapX(value);
        commands.push({
          kind: 'fill-rect',
          rect: {
            x: Math.min(zero, edge),
            y: plot.y + pointIndex * band + band * 0.14 + seriesIndex * thickness,
            width: Math.abs(edge - zero),
            height: thickness,
          },
          color: palette[seriesIndex % palette.length]!,
        });
      } else {
        const band = plot.width / pointCount;
        const thickness = (band * 0.72) / seriesCount;
        const zero = mapY(0);
        const edge = mapY(value);
        commands.push({
          kind: 'fill-rect',
          rect: {
            x: plot.x + pointIndex * band + band * 0.14 + seriesIndex * thickness,
            y: Math.min(zero, edge),
            width: thickness,
            height: Math.abs(edge - zero),
          },
          color: palette[seriesIndex % palette.length]!,
        });
      }
    });
  });
  return commands;
}

function pieCommands(model: NormalizedChart, plot: DisplayRect): readonly PrintDisplayCommand[] {
  const values = model.series[0]?.values.map((value) => Math.max(0, value ?? 0)) ?? [];
  const total = values.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return [];
  const radius = Math.max(0, Math.min(plot.width, plot.height) / 2);
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  let angle = -Math.PI / 2;
  return values.flatMap((value, index): PrintDisplayCommand[] => {
    if (value === 0) return [];
    const next = angle + (value / total) * Math.PI * 2;
    if (value === total) {
      const oppositeX = cx - Math.cos(angle) * radius;
      const oppositeY = cy - Math.sin(angle) * radius;
      const startX = cx + Math.cos(angle) * radius;
      const startY = cy + Math.sin(angle) * radius;
      angle = next;
      return [
        {
          kind: 'path',
          data: `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 1 1 ${oppositeX} ${oppositeY} A ${radius} ${radius} 0 1 1 ${startX} ${startY} Z`,
          fill: palette[index % palette.length],
        },
      ];
    }
    const x1 = cx + Math.cos(angle) * radius;
    const y1 = cy + Math.sin(angle) * radius;
    const x2 = cx + Math.cos(next) * radius;
    const y2 = cy + Math.sin(next) * radius;
    const large = next - angle > Math.PI ? 1 : 0;
    angle = next;
    return [
      {
        kind: 'path',
        data: `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z`,
        fill: palette[index % palette.length],
      },
    ];
  });
}

/** Lays out a normalized chart as bounded vector display-list operations. */
export function chartToDisplayCommands(
  model: NormalizedChart,
  rect: DisplayRect,
): ChartDisplayOutput {
  finiteRect(rect);
  const plot = plotRect(rect, model.title !== undefined);
  const body = model.type === 'pie' ? pieCommands(model, plot) : cartesianCommands(model, plot);
  const commands: readonly PrintDisplayCommand[] = Object.freeze([
    ...titleCommand(model, rect),
    {
      kind: 'clip',
      rect,
      commands: Object.freeze([...body]),
    },
  ]);
  return Object.freeze({ commands, summary: chartSummary(model) });
}
