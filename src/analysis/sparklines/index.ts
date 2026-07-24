import type { Diagnostic } from '../../document';
import type { DisplayRect, PrintDisplayCommand } from '../../print';
import {
  checkedIdentifier,
  checkedPositiveLimit,
  rangeCellCount,
  rangesIntersect,
  snapshotAddress,
  snapshotRange,
  type AnalysisCellReference,
  type AnalysisRangeReference,
} from '../references';

/** Supported compact in-cell visualization families. */
export type SparklineType = 'line' | 'column' | 'win-loss';

/** Persistent sparkline definition. The target stores presentation, not a cell value. */
export interface SparklineDefinition {
  readonly id: string;
  readonly type: SparklineType;
  readonly source: AnalysisRangeReference;
  readonly target: AnalysisCellReference;
  readonly color?: string;
  readonly negativeColor?: string;
}

/** Bounded scalar range reader evaluated against one immutable revision. */
export interface SparklineValueSource {
  readonly revision: string;
  readonly read: (reference: AnalysisRangeReference) => readonly unknown[];
}

/** Immutable sparkline snapshot shared by screen and output renderers. */
export interface NormalizedSparkline {
  readonly id: string;
  readonly type: SparklineType;
  readonly sourceRevision: string;
  readonly values: readonly (number | null)[];
  readonly target: AnalysisCellReference;
  readonly color?: string;
  readonly negativeColor?: string;
}

/** Stable result of bounded sparkline resolution. */
export interface SparklineResolution {
  readonly model: NormalizedSparkline;
  readonly dependency: AnalysisRangeReference;
  readonly diagnostics: readonly Diagnostic[];
}

/** Vector display-list commands plus screen-reader-equivalent content. */
export interface SparklineDisplayOutput {
  readonly commands: readonly PrintDisplayCommand[];
  readonly summary: string;
}

const defaultMaximumPoints = 10_000;

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

function sparklineType(value: SparklineType): SparklineType {
  if (!['line', 'column', 'win-loss'].includes(value)) {
    throw new TypeError('Sparkline type is unsupported');
  }
  return value;
}

/** Resolves a sparkline range without reading or mutating the target cell value. */
export function resolveSparkline(
  definition: SparklineDefinition,
  source: SparklineValueSource,
  limits?: { readonly maximumPoints?: number },
): SparklineResolution {
  const maximumPoints = checkedPositiveLimit(
    limits?.maximumPoints,
    defaultMaximumPoints,
    'Sparkline point limit',
  );
  const id = checkedIdentifier(definition.id, 'Sparkline ID');
  const type = sparklineType(definition.type);
  const dependency = snapshotRange(definition.source, 'Sparkline source range');
  const target = snapshotAddress(definition.target, 'Sparkline target');
  const diagnostics: Diagnostic[] = [];
  const count = rangeCellCount(dependency);
  let values: readonly (number | null)[] = [];
  if (count > maximumPoints) {
    diagnostics.push(
      diagnostic(
        'SPARKLINE_POINT_LIMIT_EXCEEDED',
        `Sparkline ${id} source exceeds the ${maximumPoints} point limit`,
        { sparklineId: id, actual: count, maximum: maximumPoints },
      ),
    );
  } else {
    const input = source.read(dependency);
    if (!Array.isArray(input)) throw new TypeError('Sparkline source must return an array');
    values = Object.freeze(
      input.slice(0, count).map((value, pointIndex) => {
        if (typeof value === 'number' && Number.isFinite(value)) return value;
        if (value === null || value === undefined || value === '') return null;
        diagnostics.push(
          diagnostic(
            'SPARKLINE_NON_NUMERIC_VALUE',
            `Sparkline ${id} point ${pointIndex} is not numeric`,
            { sparklineId: id, pointIndex },
          ),
        );
        return null;
      }),
    );
  }
  const model = Object.freeze({
    id,
    type,
    sourceRevision: source.revision,
    values,
    target,
    ...(definition.color === undefined ? {} : { color: definition.color }),
    ...(definition.negativeColor === undefined ? {} : { negativeColor: definition.negativeColor }),
  });
  return Object.freeze({
    model,
    dependency,
    diagnostics: Object.freeze(diagnostics),
  });
}

/** Returns whether a change intersects the source; the target never creates a dependency cycle. */
export function sparklineAffectedByChanges(
  definition: SparklineDefinition,
  changes: readonly AnalysisRangeReference[],
): boolean {
  return changes.some((change) => rangesIntersect(definition.source, change));
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
    throw new RangeError('Sparkline rectangle must be finite and non-negative');
  }
}

function numericBounds(values: readonly (number | null)[]): {
  readonly minimum: number;
  readonly maximum: number;
} {
  const numeric = values.filter((value): value is number => value !== null);
  if (numeric.length === 0) return { minimum: -1, maximum: 1 };
  const minimum = Math.min(0, ...numeric);
  const maximum = Math.max(0, ...numeric);
  return minimum === maximum
    ? { minimum: minimum - 1, maximum: maximum + 1 }
    : { minimum, maximum };
}

/** Lays out a normalized sparkline inside a caller-owned cell rectangle. */
export function sparklineToDisplayCommands(
  model: NormalizedSparkline,
  rect: DisplayRect,
): SparklineDisplayOutput {
  finiteRect(rect);
  const inset = 2;
  const plot = {
    x: rect.x + inset,
    y: rect.y + inset,
    width: Math.max(0, rect.width - inset * 2),
    height: Math.max(0, rect.height - inset * 2),
  };
  const { minimum, maximum } = numericBounds(model.values);
  const span = maximum - minimum;
  const baseline = plot.y + plot.height - ((0 - minimum) / span) * plot.height;
  const color = model.color ?? '#2563eb';
  const negativeColor = model.negativeColor ?? '#dc2626';
  const commands: PrintDisplayCommand[] = [];
  if (model.type === 'line') {
    const segments: string[] = [];
    let startsSegment = true;
    model.values.forEach((value, index) => {
      if (value === null) {
        startsSegment = true;
        return;
      }
      const x =
        plot.x +
        (model.values.length <= 1
          ? plot.width / 2
          : (index / (model.values.length - 1)) * plot.width);
      const y = plot.y + plot.height - ((value - minimum) / span) * plot.height;
      segments.push(`${startsSegment ? 'M' : 'L'} ${x} ${y}`);
      startsSegment = false;
      commands.push({
        kind: 'fill-rect',
        rect: { x: x - 1, y: y - 1, width: 2, height: 2 },
        color: value < 0 ? negativeColor : color,
      });
    });
    if (segments.length > 0) {
      commands.push({ kind: 'path', data: segments.join(' '), stroke: color, width: 1.5 });
    }
  } else {
    const band = plot.width / Math.max(1, model.values.length);
    model.values.forEach((value, index) => {
      if (value === null) return;
      const displayValue = model.type === 'win-loss' ? Math.sign(value) : value;
      if (displayValue === 0) return;
      const edge =
        model.type === 'win-loss'
          ? displayValue > 0
            ? plot.y
            : plot.y + plot.height
          : plot.y + plot.height - ((displayValue - minimum) / span) * plot.height;
      commands.push({
        kind: 'fill-rect',
        rect: {
          x: plot.x + index * band + band * 0.15,
          y: Math.min(baseline, edge),
          width: band * 0.7,
          height: Math.abs(edge - baseline),
        },
        color: displayValue < 0 ? negativeColor : color,
      });
    });
  }
  return Object.freeze({
    commands: Object.freeze(commands),
    summary: `Sparkline ${model.id}, ${model.type}: ${model.values
      .map((value) => value ?? 'blank')
      .join(', ')}`,
  });
}
