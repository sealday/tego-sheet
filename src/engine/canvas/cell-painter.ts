import { parseA1Range } from '../../core/coordinates/ranges';
import { getCellData } from '../../core/model/cells';
import type { CellPoint, CellRange } from '../../core/types/coordinates';
import type { CellStyle, SheetData } from '../../core/types/workbook';
import type { CellData } from '../../core/types/workbook';
import type {
  CellPresentation,
  LegacyPresentationResolver,
  ResolvedStyle,
} from '../../presentation';
import { createLegacyPresentationResolver } from '../../presentation';
import { cellRect } from '../geometry/grid-geometry';
import type { CssRect, ViewportMetrics } from '../ports';
import type { DrawContext } from './draw-context';

const DEFAULT_STYLE: CellStyle = {
  bgcolor: '#ffffff',
  align: 'left',
  valign: 'middle',
  textwrap: false,
  color: '#0a0a0a',
  font: { name: 'Arial', size: 10, bold: false, italic: false },
};

export function configuredCellDefaultStyle(configured: CellStyle | undefined): CellStyle {
  if (configured === undefined) return DEFAULT_STYLE;
  return {
    ...DEFAULT_STYLE,
    ...configured,
    font: {
      ...DEFAULT_STYLE.font,
      ...configured.font,
    },
  };
}

export interface CellPaintSnapshot {
  readonly sheet: Readonly<SheetData>;
  readonly viewport: ViewportMetrics;
  readonly invalidCells?: readonly CellPoint[];
}

/** @deprecated Use the shared presentation resolver. */
export interface LegacyCellPresentation {
  readonly cell: CellData | null;
  readonly style: CellStyle;
  readonly text: string;
  readonly printable: boolean;
}

/** @deprecated Compatibility projection for older engine tests and consumers. */
export function resolveCellPresentation(
  sheet: Readonly<SheetData>,
  point: CellPoint,
  print: boolean,
  _budget?: unknown,
  defaultStyle: CellStyle = DEFAULT_STYLE,
): LegacyCellPresentation {
  const presentation = createLegacyPresentationResolver(sheet, defaultStyle).resolve(
    point,
    print ? 'print' : 'screen',
  );
  return {
    cell: getCellData(sheet, point.row, point.column),
    style: {
      bgcolor: presentation.style.backgroundColor,
      color: presentation.style.color,
      align: presentation.style.horizontalAlign,
      valign: presentation.style.verticalAlign,
      textwrap: presentation.style.wrap,
      ...(presentation.style.numberFormat === undefined
        ? {}
        : { format: presentation.style.numberFormat }),
      font: {
        name: presentation.style.fontFamily,
        size: presentation.style.fontSize,
        bold: presentation.style.bold,
        italic: presentation.style.italic,
      },
    },
    text: presentation.formattedText,
    printable: presentation.visibility.printable,
  };
}

function isMergeAnchor(point: CellPoint, viewport: ViewportMetrics): boolean {
  const merge = viewport.model.mergeAt(point);
  return merge === null || (merge.start.row === point.row && merge.start.column === point.column);
}

const FONT_POINT_PIXELS = new Map<number, number>([
  [7.5, 10],
  [8, 11],
  [9, 12],
  [10, 13],
  [10.5, 14],
  [11, 15],
  [12, 16],
  [14, 18.7],
  [15, 20],
  [16, 21.3],
  [18, 24],
  [22, 29.3],
  [24, 32],
  [26, 34.7],
  [36, 48],
  [42, 56],
]);

function fontPixelSize(style: ResolvedStyle, visualScale: number): number {
  const pointSize = style.fontSize;
  return (FONT_POINT_PIXELS.get(pointSize) ?? pointSize) * visualScale;
}

function fontString(style: ResolvedStyle, visualScale: number): string {
  const size = fontPixelSize(style, visualScale);
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${size}px ${style.fontFamily}`;
}

function textX(rect: CssRect, style: ResolvedStyle, visualScale: number): number {
  const padding = 5 * visualScale;
  return style.horizontalAlign === 'right'
    ? rect.left + rect.width - padding
    : style.horizontalAlign === 'center'
      ? rect.left + rect.width / 2
      : rect.left + padding;
}

function border(draw: DrawContext, rect: CssRect, style: ResolvedStyle, visualScale: number): void {
  const borders = style.border;
  if (borders === undefined) return;
  const sides = [
    ['top', { x: rect.left, y: rect.top }, { x: rect.left + rect.width, y: rect.top }],
    [
      'right',
      { x: rect.left + rect.width, y: rect.top },
      { x: rect.left + rect.width, y: rect.top + rect.height },
    ],
    [
      'bottom',
      { x: rect.left, y: rect.top + rect.height },
      { x: rect.left + rect.width, y: rect.top + rect.height },
    ],
    ['left', { x: rect.left, y: rect.top }, { x: rect.left, y: rect.top + rect.height }],
  ] as const;
  for (const [side, start, end] of sides) {
    const line = borders[side];
    if (line !== undefined) {
      draw.line(start, end, {
        style: line[0],
        color: line[1] ?? '#000000',
        scale: visualScale,
      });
    }
  }
}

function marker(draw: DrawContext, rect: CssRect, color: string): void {
  const right = rect.left + rect.width - 1;
  const top = rect.top - 1;
  draw.triangle(
    [
      { x: right - 8, y: top },
      { x: right, y: top },
      { x: right, y: top + 8 },
    ],
    color,
  );
}

function dropdown(draw: DrawContext, rect: CssRect): void {
  const left = rect.left + rect.width - 15;
  const top = rect.top + rect.height - 15;
  draw.triangle(
    [
      { x: left, y: top },
      { x: left + 8, y: top },
      { x: left + 4, y: top + 6 },
    ],
    'rgba(0, 0, 0, .45)',
  );
}

function filterHeaderRange(sheet: Readonly<SheetData>): CellRange | null {
  const reference = sheet.autofilter?.ref;
  if (reference === undefined) return null;
  try {
    const range = parseA1Range(reference);
    return { start: range.start, end: { row: range.start.row, column: range.end.column } };
  } catch {
    return null;
  }
}

function pointInRange(point: CellPoint, range: CellRange | null): boolean {
  return (
    range !== null &&
    point.row >= range.start.row &&
    point.row <= range.end.row &&
    point.column >= range.start.column &&
    point.column <= range.end.column
  );
}

function wrapLines(
  draw: DrawContext,
  text: string,
  font: string,
  width: number,
  wrap: boolean,
  visualScale: number,
): readonly string[] {
  const source = text.split('\n');
  if (!wrap || width <= 0) return source;
  const output: string[] = [];
  for (const original of source) {
    if (draw.measurement.measureText(original, font) <= width) {
      output.push(original);
      continue;
    }
    let line = '';
    let lineWidth = 0;
    for (const character of original) {
      if (lineWidth >= width) {
        output.push(line);
        line = '';
        lineWidth = 0;
      }
      line += character;
      lineWidth += draw.measurement.measureText(character, font) + visualScale;
    }
    output.push(line);
  }
  return output;
}

function cellContentRect(rect: CssRect, visualScale: number): CssRect {
  return {
    left: rect.left + visualScale,
    top: rect.top + visualScale,
    width: Math.max(0, rect.width - 2 * visualScale),
    height: Math.max(0, rect.height - 2 * visualScale),
  };
}

function paintCellContent(
  draw: DrawContext,
  rect: CssRect,
  presentation: CellPresentation,
  visualScale: number,
): void {
  const style = presentation.style;
  draw.fillRect(cellContentRect(rect, visualScale), style.backgroundColor);
  if (presentation.formattedText === '') return;
  const font = fontString(style, visualScale);
  const lines = wrapLines(
    draw,
    presentation.formattedText,
    font,
    Math.max(0, rect.width - 12 * visualScale),
    style.wrap,
    visualScale,
  );
  const x = textX(rect, style, visualScale);
  const fontSize = fontPixelSize(style, visualScale);
  const lineHeight = fontSize + 2 * visualScale;
  const textHeight = (lines.length - 1) * lineHeight;
  const padding = 5 * visualScale;
  const firstY =
    style.verticalAlign === 'top'
      ? rect.top + padding
      : style.verticalAlign === 'bottom'
        ? rect.top + rect.height - padding - textHeight
        : rect.top + rect.height / 2 - textHeight / 2;
  for (const [index, line] of lines.entries()) {
    const y = firstY + index * lineHeight;
    draw.text(
      line,
      { x, y },
      {
        align: style.horizontalAlign,
        baseline: style.verticalAlign,
        color: style.color,
        font,
      },
    );
    const width = draw.measurement.measureText(line, font);
    const startX =
      style.horizontalAlign === 'center'
        ? x - width / 2
        : style.horizontalAlign === 'right'
          ? x - width
          : x;
    if (style.underline === true) {
      const underlineY =
        style.verticalAlign === 'top'
          ? y + fontSize + 2 * visualScale
          : style.verticalAlign === 'bottom'
            ? y
            : y + fontSize / 2;
      draw.line(
        { x: startX, y: underlineY },
        { x: startX + width, y: underlineY },
        { color: style.color, scale: visualScale },
      );
    }
    if (style.strike === true) {
      const strikeY =
        style.verticalAlign === 'top'
          ? y + fontSize / 2 + 2 * visualScale
          : style.verticalAlign === 'bottom'
            ? y - fontSize / 2
            : y;
      draw.line(
        { x: startX, y: strikeY },
        { x: startX + width, y: strikeY },
        {
          color: style.color,
          scale: visualScale,
        },
      );
    }
  }
}

export function paintCellAppearance(
  draw: DrawContext,
  rect: CssRect,
  presentation: CellPresentation,
  visualScale = 1,
  paintMarks?: () => void,
): void {
  const style = presentation.style;
  border(draw, rect, style, visualScale);
  draw.withPixelAlignedClip(cellContentRect(rect, visualScale), () => {
    paintCellContent(draw, rect, presentation, visualScale);
    paintMarks?.();
  });
}

export function paintCells(
  draw: DrawContext,
  snapshot: CellPaintSnapshot,
  cells: readonly CellPoint[],
  presentations: LegacyPresentationResolver,
): void {
  const invalid = new Set(
    (snapshot.invalidCells ?? []).map((point) => `${point.row}:${point.column}`),
  );
  for (const point of cells) {
    if (getCellData(snapshot.sheet, point.row, point.column) === null) continue;
    if (!isMergeAnchor(point, snapshot.viewport)) continue;
    const rect = cellRect(point, snapshot.viewport);
    const presentation = presentations.resolve(point, 'screen');
    paintCellAppearance(draw, rect, presentation, 1, () => {
      if (invalid.has(`${point.row}:${point.column}`) || presentation.accessibility.invalid)
        marker(draw, rect, 'rgba(255, 0, 0, .65)');
      if (presentation.accessibility.readOnly) marker(draw, rect, 'rgba(0, 255, 0, .85)');
    });
  }
}

export function paintFilterOverlays(
  draw: DrawContext,
  snapshot: CellPaintSnapshot,
  visibleRows: readonly number[],
  visibleColumns: readonly number[],
): void {
  const filter = filterHeaderRange(snapshot.sheet);
  if (filter === null) return;
  for (const row of visibleRows) {
    if (row < filter.start.row || row > filter.end.row) continue;
    for (const column of visibleColumns) {
      const point = { row, column };
      if (pointInRange(point, filter)) dropdown(draw, cellRect(point, snapshot.viewport));
    }
  }
}
