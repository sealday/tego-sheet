import type {
  Diagnostic,
  DocumentCellRange,
  DocumentSheetId,
  Sheet,
  SpreadsheetDocument,
} from '../document';
import { createFormulaEngine } from '../formula';
import {
  createPresentationCache,
  createPresentationResolver,
  type CellPresentation,
} from '../presentation';
import {
  createPrintDisplayList,
  type PrintDisplayCommand,
  type PrintDisplayPageInput,
} from '../print';
import {
  paginateTemplateTargets,
  type PaginationPage,
  type PaginationTarget,
} from '../print/layout';
import {
  expandAdvancedTemplate,
  expandTemplate,
  type RowInsertion,
  type StructuralMapping,
  type StructuralObjectMapping,
} from './expand';
import { TemplateExpressionError } from './expression';
import {
  createResourceResolverRegistry,
  resolveTemplateResources,
  type ResolvedResourceStore,
} from './resources';
import type {
  GeneratedDocument,
  RenderEnvironment,
  RenderLimits,
  RenderRequest,
  RenderResult,
  TemplatePrintProfile,
  TemplateResourceBinding,
} from './model';
import { objectToDisplayCommands, resolveObjectAnchor } from '../objects';
import { applyDocumentFilterView } from '../views';
import {
  createPersistedVisualizationValueSource,
  projectPersistedVisualizations,
  type ChartValueSource,
} from '../analysis';

const DEFAULT_LIMITS: RenderLimits = Object.freeze({
  maxExpandedCells: 250_000,
  maxExpandedRows: 100_000,
  maxPages: 10_000,
  maxLayoutTimeMs: 5_000,
  maxExpandedColumns: 16_384,
  maxGeneratedSheets: 256,
  maxExpandedObjects: 10_000,
  maxNestingDepth: 8,
  maxResources: 256,
  maxResourceBytes: 16 * 1024 * 1024,
  maxTotalResourceBytes: 64 * 1024 * 1024,
  maxResolveConcurrency: 4,
});

const PAPER = Object.freeze({
  A4: Object.freeze({ width: 794, height: 1123 }),
  A5: Object.freeze({ width: 559, height: 794 }),
  Letter: Object.freeze({ width: 816, height: 1056 }),
});

interface ResolvedTarget {
  readonly id: string;
  readonly sheet: Sheet;
  readonly range: DocumentCellRange;
}

interface RenderExpansion {
  readonly document?: SpreadsheetDocument;
  readonly diagnostics: readonly Diagnostic[];
  readonly insertedRows: ReadonlyMap<string, readonly RowInsertion[]>;
  readonly repeatPageBreaks: ReadonlyMap<string, readonly number[]>;
  readonly structuralMappings: readonly StructuralMapping[];
  readonly objectMappings: readonly StructuralObjectMapping[];
}

function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (ArrayBuffer.isView(value) || value instanceof Map || value instanceof Set) return value;
  if (value !== null && typeof value === 'object') {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

function renderDiagnostic(code: string, message: string): Diagnostic {
  return freeze({
    code,
    severity: 'error',
    domain: 'template',
    stage: code === 'TEMPLATE_SOURCE_STALE' ? 'validate' : 'render',
    message,
  });
}

function mappedRow(row: number, insertions: readonly RowInsertion[]): number {
  return (
    row +
    insertions.reduce(
      (delta, insertion) => delta + (row > insertion.afterSourceRow ? insertion.delta : 0),
      0,
    )
  );
}

function usedRange(sheet: Sheet): DocumentCellRange {
  const maximumRow = Math.max(
    0,
    (sheet.rowCount ?? 1) - 1,
    ...sheet.cells.map(({ row }) => row),
    ...sheet.rows.map(({ index }) => index),
  );
  const maximumColumn = Math.max(
    0,
    (sheet.columnCount ?? 1) - 1,
    ...sheet.cells.map(({ column }) => column),
    ...sheet.columns.map(({ index }) => index),
  );
  return {
    sheetId: sheet.id,
    start: { row: 0, column: 0 },
    end: { row: maximumRow, column: maximumColumn },
  };
}

function transformRange(
  range: DocumentCellRange,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  structuralMappings: readonly StructuralMapping[] = [],
): DocumentCellRange {
  const sheetInsertions = insertions.get(range.sheetId) ?? [];
  const mapped = {
    ...range,
    start: { ...range.start, row: mappedRow(range.start.row, sheetInsertions) },
    end: {
      ...range.end,
      row:
        range.end.row +
        sheetInsertions.reduce(
          (delta, insertion) =>
            delta + (range.end.row >= insertion.afterSourceRow ? insertion.delta : 0),
          0,
        ),
    },
  };
  const relevant = structuralMappings.filter(
    ({ source }) =>
      source.sheetId === range.sheetId &&
      source.start.row >= range.start.row &&
      source.end.row <= range.end.row &&
      source.start.column >= range.start.column &&
      source.end.column <= range.end.column,
  );
  const rowDelta = relevant.reduce(
    (maximum, { source, generated }) => Math.max(maximum, generated.end.row - source.end.row),
    0,
  );
  const columnDelta = relevant.reduce(
    (maximum, { source, generated }) => Math.max(maximum, generated.end.column - source.end.column),
    0,
  );
  return {
    ...mapped,
    end: {
      row: mapped.end.row + rowDelta,
      column: mapped.end.column + columnDelta,
    },
  };
}

function generatedSheetIds(
  sourceSheetId: string,
  structuralMappings: readonly StructuralMapping[],
): readonly string[] {
  return [
    ...new Set(
      structuralMappings
        .filter(
          ({ source, generated }) =>
            source.sheetId === sourceSheetId && generated.sheetId !== sourceSheetId,
        )
        .map(({ generated }) => String(generated.sheetId)),
    ),
  ];
}

function profileRangeForSheet(
  range: DocumentCellRange,
  sheetId: string,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  structuralMappings: readonly StructuralMapping[],
): DocumentCellRange | undefined {
  if (
    range.sheetId !== sheetId &&
    !structuralMappings.some(
      ({ source, generated }) => source.sheetId === range.sheetId && generated.sheetId === sheetId,
    )
  ) {
    return undefined;
  }
  return {
    ...transformRange(range, insertions, structuralMappings),
    sheetId: sheetId as DocumentCellRange['sheetId'],
  };
}

function targets(
  document: SpreadsheetDocument,
  profile: TemplatePrintProfile,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  structuralMappings: readonly StructuralMapping[] = [],
): readonly ResolvedTarget[] {
  const output: ResolvedTarget[] = [];
  const append = (range: DocumentCellRange, transform = true): void => {
    const transformed = transform ? transformRange(range, insertions, structuralMappings) : range;
    const sheet = document.workbook.sheets.find(({ id }) => id === transformed.sheetId);
    if (sheet === undefined) return;
    output.push({
      id: `${profile.id}:${output.length}:${sheet.id}:${transformed.start.row}-${transformed.end.row}:${transformed.start.column}-${transformed.end.column}`,
      sheet,
      range: transformed,
    });
  };
  const appendGeneratedRanges = (range: DocumentCellRange): boolean => {
    const generated = generatedSheetIds(range.sheetId, structuralMappings);
    if (generated.length === 0) return false;
    const transformed = transformRange(range, insertions, structuralMappings);
    generated.forEach((sheetId) =>
      append({ ...transformed, sheetId: sheetId as DocumentCellRange['sheetId'] }, false),
    );
    return true;
  };
  for (const target of profile.targets) {
    if (target.type === 'sheet') {
      const generated = generatedSheetIds(target.sheetId, structuralMappings);
      if (generated.length > 0) {
        for (const sheetId of generated) {
          const generatedSheet = document.workbook.sheets.find(({ id }) => id === sheetId);
          if (generatedSheet !== undefined) append(usedRange(generatedSheet), false);
        }
        continue;
      }
      const sheet = document.workbook.sheets.find(({ id }) => id === target.sheetId);
      if (sheet !== undefined) append(usedRange(sheet), false);
    } else if (target.type === 'range') {
      if (!appendGeneratedRanges(target.range)) append(target.range);
    } else {
      target.ranges.forEach((range) => {
        if (!appendGeneratedRanges(range)) append(range);
      });
    }
  }
  return output;
}

function rowHeight(sheet: Sheet, row: number): number {
  return sheet.rows.find(({ index }) => index === row)?.height ?? 20;
}

function columnWidth(sheet: Sheet, column: number): number {
  return sheet.columns.find(({ index }) => index === column)?.width ?? 100;
}

function paper(profile: TemplatePrintProfile): { readonly width: number; readonly height: number } {
  const definition =
    profile.page.paper.type === 'custom'
      ? { width: profile.page.paper.width, height: profile.page.paper.height }
      : PAPER[profile.page.paper.type];
  return profile.page.orientation === 'portrait'
    ? definition
    : { width: definition.height, height: definition.width };
}

interface ViewProjection {
  readonly hiddenRows: ReadonlySet<number>;
  readonly rowOrder: readonly number[];
  readonly startRow: number;
  readonly endRow: number;
}

function projectedTargetRows(
  range: DocumentCellRange,
  projection: ViewProjection | undefined,
): readonly number[] {
  const rows = Array.from(
    { length: range.end.row - range.start.row + 1 },
    (_, index) => range.start.row + index,
  );
  if (projection === undefined) return rows;
  return [
    ...rows.filter((row) => row < projection.startRow),
    ...projection.rowOrder.filter(
      (row) => row >= range.start.row && row <= range.end.row && !projection.hiddenRows.has(row),
    ),
    ...rows.filter((row) => row > projection.endRow),
  ];
}

function projectedBreakIndex(
  target: ResolvedTarget,
  logicalRow: number,
  projection: ViewProjection | undefined,
): number | undefined {
  const index = projectedTargetRows(target.range, projection).indexOf(logicalRow);
  return index < 0 ? undefined : index;
}

interface ProjectedRowGeometry {
  readonly rows: readonly number[];
  readonly offsets: readonly number[];
  markerOffset(row: number): number | undefined;
}

/** Builds the global visual row geometry shared by every page and object for one print target. */
function createProjectedRowGeometry(
  sheet: Sheet,
  range: DocumentCellRange,
  projection: ViewProjection | undefined,
): ProjectedRowGeometry {
  const rows = projectedTargetRows(range, projection);
  const rangeOrigin = Array.from({ length: range.start.row }, (_, index) =>
    rowHeight(sheet, index),
  ).reduce((sum, value) => sum + value, 0);
  const offsets = [rangeOrigin];
  const markerOffsets = new Map<number, number>();
  for (const row of rows) {
    markerOffsets.set(row, offsets.at(-1)!);
    offsets.push(offsets.at(-1)! + rowHeight(sheet, row));
  }
  return {
    rows,
    offsets,
    markerOffset(row) {
      if (row < range.start.row) {
        return Array.from({ length: row }, (_, index) => rowHeight(sheet, index)).reduce(
          (sum, value) => sum + value,
          0,
        );
      }
      if (row > range.end.row) {
        const trailingExtent = Array.from(
          { length: row - range.end.row - 1 },
          (_, index) => range.end.row + 1 + index,
        ).reduce((sum, candidate) => sum + rowHeight(sheet, candidate), 0);
        return offsets.at(-1)! + trailingExtent;
      }
      return markerOffsets.get(row);
    },
  };
}

function paginationTargets(
  resolved: readonly ResolvedTarget[],
  profile: TemplatePrintProfile,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  structuralMappings: readonly StructuralMapping[] = [],
  viewProjections: ReadonlyMap<string, ViewProjection> = new Map(),
): readonly PaginationTarget[] {
  return resolved.map(({ id, sheet, range }) => {
    const repeatedRows =
      profile.repeatRows === undefined
        ? undefined
        : profileRangeForSheet(profile.repeatRows, sheet.id, insertions, structuralMappings);
    const repeatedColumns =
      profile.repeatColumns === undefined
        ? undefined
        : profileRangeForSheet(profile.repeatColumns, sheet.id, insertions, structuralMappings);
    return {
      id,
      rows: projectedTargetRows(range, viewProjections.get(sheet.id)).map((row) =>
        rowHeight(sheet, row),
      ),
      columns: Array.from({ length: range.end.column - range.start.column + 1 }, (_, index) =>
        columnWidth(sheet, range.start.column + index),
      ),
      ...(repeatedRows === undefined
        ? {}
        : {
            repeatRows: Array.from(
              { length: repeatedRows.end.row - repeatedRows.start.row + 1 },
              (_, index) => rowHeight(sheet, repeatedRows.start.row + index),
            ),
          }),
      ...(repeatedColumns === undefined
        ? {}
        : {
            repeatColumns: Array.from(
              { length: repeatedColumns.end.column - repeatedColumns.start.column + 1 },
              (_, index) => columnWidth(sheet, repeatedColumns.start.column + index),
            ),
          }),
    };
  });
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

function bandText(source: string, page: number, pages: number, data: unknown, date: Date): string {
  return source.replace(/\{\{\s*([^}]+?)\s*\}\}/gu, (_, token: string) => {
    if (token === 'page') return String(page);
    if (token === 'pages') return String(pages);
    if (token === 'date') return date.toISOString().slice(0, 10);
    let value = data;
    for (const part of token.split('.')) {
      if (
        ['__proto__', 'prototype', 'constructor'].includes(part) ||
        value === null ||
        typeof value !== 'object' ||
        !Object.prototype.hasOwnProperty.call(value, part)
      ) {
        return '';
      }
      value = (value as Readonly<Record<string, unknown>>)[part];
    }
    return value === undefined || value === null ? '' : String(value);
  });
}

function bandCommands(
  profile: TemplatePrintProfile,
  page: PaginationPage,
  totalPages: number,
  data: unknown,
  date: Date,
): readonly PrintDisplayCommand[] {
  const commands: PrintDisplayCommand[] = [];
  for (const [band, y] of [
    [profile.header, Math.max(10, profile.page.margins.top / 2)],
    [profile.footer, page.height - Math.max(10, profile.page.margins.bottom / 2)],
  ] as const) {
    if (band === undefined) continue;
    for (const [source, align, x] of [
      [band.left, 'left', profile.page.margins.left],
      [band.center, 'center', page.width / 2],
      [band.right, 'right', page.width - profile.page.margins.right],
    ] as const) {
      if (source === undefined) continue;
      commands.push({
        kind: 'text',
        text: bandText(source, page.index + 1, totalPages, data, date),
        x,
        y,
        maxWidth: Math.max(0, page.width / 3),
        fontFamily: 'Arial',
        fontSize: 10,
        color: '#0a0a0a',
        horizontalAlign: align,
      });
    }
  }
  return commands;
}

function rotatedObjectBounds(
  rect: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  rotation: number | undefined,
  padding: number,
): { readonly x: number; readonly y: number; readonly width: number; readonly height: number } {
  const angle = ((rotation ?? 0) * Math.PI) / 180;
  const cosine = Math.abs(Math.cos(angle));
  const sine = Math.abs(Math.sin(angle));
  const width = rect.width * cosine + rect.height * sine + padding * 2;
  const height = rect.width * sine + rect.height * cosine + padding * 2;
  return {
    x: rect.x + rect.width / 2 - width / 2,
    y: rect.y + rect.height / 2 - height / 2,
    width,
    height,
  };
}

function displayPages(
  pages: readonly PaginationPage[],
  resolved: readonly ResolvedTarget[],
  profile: TemplatePrintProfile,
  resolvePresentation: (sheetId: DocumentSheetId, row: number, column: number) => CellPresentation,
  insertions: ReadonlyMap<string, readonly RowInsertion[]>,
  structuralMappings: readonly StructuralMapping[],
  resourceBindings: readonly TemplateResourceBinding[],
  resources: ResolvedResourceStore,
  viewProjections: ReadonlyMap<string, ViewProjection>,
  data: unknown,
  date: Date,
  visualizationSource: ChartValueSource,
): readonly PrintDisplayPageInput[] {
  const projectedGeometry = new Map(
    resolved.map((target) => [
      target.id,
      createProjectedRowGeometry(target.sheet, target.range, viewProjections.get(target.sheet.id)),
    ]),
  );
  return pages.map((page) => {
    const target = resolved.find(({ id }) => id === page.targetId)!;
    const startColumn = target.range.start.column + page.columnStart;
    const endColumn = target.range.start.column + page.columnEnd;
    const viewProjection = viewProjections.get(target.sheet.id);
    const rowGeometry = projectedGeometry.get(target.id)!;
    const projectedRows = rowGeometry.rows;
    const bodyRows = projectedRows.slice(page.rowStart, page.rowEnd + 1);
    const startRow = bodyRows.length === 0 ? target.range.start.row : Math.min(...bodyRows);
    const endRow = bodyRows.length === 0 ? target.range.start.row : Math.max(...bodyRows);
    const bodyColumns = Array.from(
      { length: endColumn - startColumn + 1 },
      (_, index) => startColumn + index,
    );
    const titleRows =
      profile.repeatRows === undefined
        ? []
        : (() => {
            const range = profileRangeForSheet(
              profile.repeatRows,
              target.sheet.id,
              insertions,
              structuralMappings,
            );
            if (range === undefined) return [];
            return Array.from(
              { length: range.end.row - range.start.row + 1 },
              (_, index) => range.start.row + index,
            ).filter((row) => !bodyRows.includes(row));
          })();
    const titleColumns =
      profile.repeatColumns === undefined
        ? []
        : (() => {
            const range = profileRangeForSheet(
              profile.repeatColumns,
              target.sheet.id,
              insertions,
              structuralMappings,
            );
            if (range === undefined) return [];
            return Array.from(
              { length: range.end.column - range.start.column + 1 },
              (_, index) => range.start.column + index,
            ).filter((column) => !bodyColumns.includes(column));
          })();
    const rows = [...titleRows, ...bodyRows];
    const columns = [...titleColumns, ...bodyColumns];
    const headingSize = profile.showHeadings ? 20 : 0;
    const cells = [];
    for (const [rowIndex, row] of rows.entries()) {
      for (const [columnIndex, column] of columns.entries()) {
        const merge = target.sheet.merges.find(
          ({ start, end }) =>
            row >= start.row && row <= end.row && column >= start.column && column <= end.column,
        );
        const mergeRows =
          merge === undefined
            ? [row]
            : rows.filter(
                (candidate) => candidate >= merge.start.row && candidate <= merge.end.row,
              );
        const mergeColumns =
          merge === undefined
            ? [column]
            : columns.filter(
                (candidate) => candidate >= merge.start.column && candidate <= merge.end.column,
              );
        const mergeOrigin =
          merge === undefined
            ? { row, column }
            : { row: mergeRows[0] ?? row, column: mergeColumns[0] ?? column };
        if (row !== mergeOrigin.row || column !== mergeOrigin.column) continue;
        cells.push({
          rect: {
            x:
              profile.page.margins.left +
              headingSize +
              columns
                .slice(0, columnIndex)
                .reduce((sum, candidate) => sum + columnWidth(target.sheet, candidate), 0) *
                page.scale,
            y:
              profile.page.margins.top +
              headingSize +
              rows
                .slice(0, rowIndex)
                .reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
                page.scale,
            width:
              mergeColumns.reduce(
                (sum, candidate) => sum + columnWidth(target.sheet, candidate),
                0,
              ) * page.scale,
            height:
              mergeRows.reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
              page.scale,
          },
          presentation: resolvePresentation(
            target.sheet.id,
            merge?.start.row ?? row,
            merge?.start.column ?? column,
          ),
        });
      }
    }
    const decorations: PrintDisplayCommand[] = [
      ...bandCommands(profile, page, pages.length, data, date),
    ];
    if (profile.showHeadings) {
      for (const [index, column] of columns.entries()) {
        decorations.push({
          kind: 'text',
          text: columnLabel(column),
          x:
            profile.page.margins.left +
            headingSize +
            columns
              .slice(0, index)
              .reduce((sum, candidate) => sum + columnWidth(target.sheet, candidate), 0) *
              page.scale +
            (columnWidth(target.sheet, column) * page.scale) / 2,
          y: profile.page.margins.top + headingSize / 2,
          maxWidth: columnWidth(target.sheet, column) * page.scale,
          fontFamily: 'Arial',
          fontSize: 10,
          color: '#0a0a0a',
          horizontalAlign: 'center',
        });
      }
      for (const [index, row] of rows.entries()) {
        decorations.push({
          kind: 'text',
          text: String(row + 1),
          x: profile.page.margins.left + headingSize / 2,
          y:
            profile.page.margins.top +
            headingSize +
            rows
              .slice(0, index)
              .reduce((sum, candidate) => sum + rowHeight(target.sheet, candidate), 0) *
              page.scale +
            (rowHeight(target.sheet, row) * page.scale) / 2,
          maxWidth: headingSize,
          fontFamily: 'Arial',
          fontSize: 10,
          color: '#0a0a0a',
          horizontalAlign: 'center',
        });
      }
    }
    const overlays = resourceBindings.flatMap((binding): readonly PrintDisplayCommand[] => {
      const placement = transformRange(binding.target, insertions, structuralMappings);
      if (placement.sheetId !== target.sheet.id) return [];
      const firstRow = Math.max(placement.start.row, startRow);
      const lastRow = Math.min(placement.end.row, endRow);
      const firstColumn = Math.max(placement.start.column, startColumn);
      const lastColumn = Math.min(placement.end.column, endColumn);
      if (firstRow > lastRow || firstColumn > lastColumn) return [];
      const rowIndex = rows.indexOf(firstRow);
      const columnIndex = columns.indexOf(firstColumn);
      if (rowIndex < 0 || columnIndex < 0) return [];
      return [
        {
          kind: 'image',
          resourceId: binding.resourceId,
          fit: binding.fit,
          rect: {
            x:
              profile.page.margins.left +
              headingSize +
              columns
                .slice(0, columnIndex)
                .reduce((sum, column) => sum + columnWidth(target.sheet, column), 0) *
                page.scale,
            y:
              profile.page.margins.top +
              headingSize +
              rows.slice(0, rowIndex).reduce((sum, row) => sum + rowHeight(target.sheet, row), 0) *
                page.scale,
            width:
              columns
                .filter((column) => column >= firstColumn && column <= lastColumn)
                .reduce((sum, column) => sum + columnWidth(target.sheet, column), 0) * page.scale,
            height:
              rows
                .filter((row) => row >= firstRow && row <= lastRow)
                .reduce((sum, row) => sum + rowHeight(target.sheet, row), 0) * page.scale,
          },
        },
      ];
    });
    const physicalRowOffset = (row: number): number =>
      Array.from({ length: row }, (_, index) => rowHeight(target.sheet, index)).reduce(
        (sum, value) => sum + value,
        0,
      );
    const columnOffset = (column: number): number =>
      Array.from({ length: column }, (_, index) => columnWidth(target.sheet, index)).reduce(
        (sum, value) => sum + value,
        0,
      );
    const bodyOriginX = columnOffset(startColumn);
    const bodyOriginY = rowGeometry.offsets[page.rowStart]!;
    const bodyBottom = rowGeometry.offsets[page.rowEnd + 1]!;
    const objectOverlays = [...(target.sheet.objects ?? [])]
      .sort((left, right) => left.zIndex - right.zIndex || left.id.localeCompare(right.id))
      .flatMap((object): readonly PrintDisplayCommand[] => {
        const primaryRow =
          object.anchor.type === 'absolute'
            ? undefined
            : object.anchor.type === 'one-cell'
              ? object.anchor.cell.row
              : object.anchor.from.row;
        const primaryRowOffset =
          primaryRow === undefined ? undefined : rowGeometry.markerOffset(primaryRow);
        // A filtered primary anchor has no visual row, so the complete object is omitted.
        if (primaryRow !== undefined && primaryRowOffset === undefined) return [];
        const rowOffset = (row: number): number =>
          rowGeometry.markerOffset(row) ?? primaryRowOffset ?? physicalRowOffset(row);
        const geometry = { rowOffset, columnOffset };
        const resolvedRect = resolveObjectAnchor(object.anchor, geometry);
        const rect =
          object.anchor.type === 'two-cell' && viewProjection !== undefined
            ? (() => {
                const fromY = rowOffset(object.anchor.from.row) + object.anchor.from.offset.y;
                const toY = rowOffset(object.anchor.to.row) + object.anchor.to.offset.y;
                return {
                  ...resolvedRect,
                  y: Math.min(fromY, toY),
                  height: Math.abs(toY - fromY),
                };
              })()
            : resolvedRect;
        const bodyRight = columnOffset(endColumn + 1);
        const strokePadding = object.kind === 'shape' ? (object.style.strokeWidth ?? 1) / 2 : 0;
        const visibleBounds = rotatedObjectBounds(rect, object.rotation, strokePadding);
        if (
          visibleBounds.x + visibleBounds.width <= bodyOriginX ||
          visibleBounds.y + visibleBounds.height <= bodyOriginY ||
          visibleBounds.x >= bodyRight ||
          visibleBounds.y >= bodyBottom
        ) {
          return [];
        }
        const pageRect = {
          x: profile.page.margins.left + headingSize + (rect.x - bodyOriginX) * page.scale,
          y: profile.page.margins.top + headingSize + (rect.y - bodyOriginY) * page.scale,
          width: rect.width * page.scale,
          height: rect.height * page.scale,
        };
        const displayObject =
          object.kind === 'text-box'
            ? {
                ...object,
                anchor: { type: 'absolute' as const, rect: pageRect },
                style: { ...object.style, fontSize: object.style.fontSize * page.scale },
              }
            : object.kind === 'shape'
              ? {
                  ...object,
                  anchor: { type: 'absolute' as const, rect: pageRect },
                  style: {
                    ...object.style,
                    ...(object.style.strokeWidth === undefined
                      ? {}
                      : { strokeWidth: object.style.strokeWidth * page.scale }),
                  },
                }
              : {
                  ...object,
                  anchor: { type: 'absolute' as const, rect: pageRect },
                };
        return objectToDisplayCommands(displayObject, {
          geometry,
          resources: resources.byReference,
        });
      });
    const visualizationOverlays = projectPersistedVisualizations(
      target.sheet,
      visualizationSource,
      {
        chart(definition) {
          if (definition.anchor === undefined) return null;
          const rect = resolveObjectAnchor(definition.anchor, {
            rowOffset: (row) => rowGeometry.markerOffset(row) ?? physicalRowOffset(row),
            columnOffset,
          });
          const bodyRight = columnOffset(endColumn + 1);
          if (
            rect.x + rect.width <= bodyOriginX ||
            rect.y + rect.height <= bodyOriginY ||
            rect.x >= bodyRight ||
            rect.y >= bodyBottom
          )
            return null;
          return {
            x: profile.page.margins.left + headingSize + (rect.x - bodyOriginX) * page.scale,
            y: profile.page.margins.top + headingSize + (rect.y - bodyOriginY) * page.scale,
            width: rect.width * page.scale,
            height: rect.height * page.scale,
          };
        },
        sparkline(definition) {
          const rowIndex = rows.indexOf(definition.target.row);
          const columnIndex = columns.indexOf(definition.target.column);
          if (rowIndex < 0 || columnIndex < 0) return null;
          return {
            x:
              profile.page.margins.left +
              headingSize +
              columns
                .slice(0, columnIndex)
                .reduce((sum, column) => sum + columnWidth(target.sheet, column), 0) *
                page.scale,
            y:
              profile.page.margins.top +
              headingSize +
              rows.slice(0, rowIndex).reduce((sum, row) => sum + rowHeight(target.sheet, row), 0) *
                page.scale,
            width: columnWidth(target.sheet, definition.target.column) * page.scale,
            height: rowHeight(target.sheet, definition.target.row) * page.scale,
          };
        },
      },
    ).flatMap(({ commands }) => commands);
    return {
      width: page.width,
      height: page.height,
      cells,
      decorations,
      overlays: [...overlays, ...objectOverlays, ...visualizationOverlays],
      showGridlines: profile.showGridlines,
    };
  });
}

function persistentImageMappings(
  document: SpreadsheetDocument,
): readonly import('./expand').StructuralObjectMapping[] {
  return document.workbook.sheets.flatMap((sheet) =>
    (sheet.objects ?? []).flatMap((object) => {
      if (object.kind !== 'image' || object.anchor.type !== 'two-cell') return [];
      const range = {
        sheetId: sheet.id,
        start: {
          row: object.anchor.from.row,
          column: object.anchor.from.column,
        },
        end: {
          row: object.anchor.to.row,
          column: object.anchor.to.column,
        },
      };
      return [
        {
          objectId: object.id,
          resourceId: object.resourceId,
          policy: 'shared' as const,
          itemIndex: 0,
          source: range,
          generated: range,
        },
      ];
    }),
  );
}

function generatedObjectMappings(
  document: SpreadsheetDocument,
  repeated: readonly StructuralObjectMapping[],
): readonly StructuralObjectMapping[] {
  const repeatedIds = new Set(repeated.map(({ objectId }) => String(objectId)));
  return [
    ...repeated,
    ...persistentImageMappings(document).filter(
      ({ objectId }) => !repeatedIds.has(String(objectId)),
    ),
  ];
}

function abortResult(): RenderResult {
  const diagnostic = renderDiagnostic('RENDER_ABORTED', 'Template rendering was aborted');
  return freeze({ diagnostics: [diagnostic] });
}

function isAborted(signal: AbortSignal | undefined): boolean {
  return signal?.aborted === true;
}

/** Executes expand → recalculate → present → paginate against one immutable source. */
export async function renderSpreadsheetTemplate(
  request: RenderRequest,
  environment: RenderEnvironment,
): Promise<RenderResult> {
  if (isAborted(request.signal)) return abortResult();
  if (request.template.sourceDocumentHash !== request.currentDocumentHash) {
    return freeze({
      diagnostics: [
        renderDiagnostic(
          'TEMPLATE_SOURCE_STALE',
          'Compiled template does not match the current document',
        ),
      ],
    });
  }
  const limits = Object.freeze({ ...DEFAULT_LIMITS, ...request.limits });
  const start = Date.now();
  const resourceController = new AbortController();
  const abortResources = (): void => resourceController.abort();
  request.signal?.addEventListener('abort', abortResources, { once: true });
  const resourceResult = await resolveTemplateResources(request.resourceRefs ?? [], {
    registry: environment.resourceRegistry ?? createResourceResolverRegistry([]),
    signal: resourceController.signal,
    purpose: environment.resourcePurpose ?? 'preview',
    limits: {
      maxResources: limits.maxResources,
      maxResourceBytes: limits.maxResourceBytes,
      maxTotalResourceBytes: limits.maxTotalResourceBytes,
      maxResolveConcurrency: limits.maxResolveConcurrency,
    },
    ...(environment.decodeImage === undefined ? {} : { decodeImage: environment.decodeImage }),
  });
  request.signal?.removeEventListener('abort', abortResources);
  if (resourceResult.store === undefined) {
    return freeze({ diagnostics: resourceResult.diagnostics });
  }
  const resources: ResolvedResourceStore = resourceResult.store;
  const failAfterResources = async (result: RenderResult): Promise<RenderResult> => {
    await resources?.dispose();
    return result;
  };
  let transferred = false;
  try {
    const profile = request.template.ir.profiles.find(({ id }) => id === request.profileId);
    if (profile === undefined) {
      return failAfterResources(
        freeze({
          diagnostics: [
            renderDiagnostic('INVALID_PRINT_TARGET', `Unknown print profile: ${request.profileId}`),
          ],
        }),
      );
    }
    const resourceBindings = request.template.ir.template.resourceBindings ?? [];
    const unresolvedResource = resourceBindings.find(
      ({ resourceId }) => resources.byReference[resourceId] === undefined,
    );
    if (unresolvedResource !== undefined) {
      return failAfterResources(
        freeze({
          diagnostics: [
            renderDiagnostic(
              'RESOURCE_BINDING_UNRESOLVED',
              `Template resource ${unresolvedResource.resourceId} was not resolved`,
            ),
          ],
        }),
      );
    }
    const unsupportedResource = resourceBindings.find(({ resourceId }) => {
      const resource = resources.byReference[resourceId];
      return resource?.type !== 'image' && resource?.type !== 'svg' && resource?.type !== 'qr';
    });
    if (unsupportedResource !== undefined) {
      return failAfterResources(
        freeze({
          diagnostics: [
            renderDiagnostic(
              'RESOURCE_TYPE_MISMATCH',
              `Template resource ${unsupportedResource.resourceId} is not displayable`,
            ),
          ],
        }),
      );
    }
    const unresolvedObject = request.template.sourceDocument.workbook.sheets
      .flatMap((sheet) => sheet.objects ?? [])
      .find(
        (object) =>
          object.kind === 'image' && resources.byReference[object.resourceId] === undefined,
      );
    if (unresolvedObject?.kind === 'image') {
      return failAfterResources(
        freeze({
          diagnostics: [
            renderDiagnostic(
              'RESOURCE_BINDING_UNRESOLVED',
              `Sheet object resource ${unresolvedObject.resourceId} was not resolved`,
            ),
          ],
        }),
      );
    }
    let expansion: RenderExpansion;
    try {
      const hasAdvanced =
        request.template.ir.regionTree !== undefined ||
        request.template.ir.bindings.some(
          (binding) =>
            (binding.type === 'repeat-rows' &&
              binding.objects !== undefined &&
              binding.objects.length > 0) ||
            binding.type === 'repeat-columns' ||
            binding.type === 'repeat-range' ||
            binding.type === 'repeat-page' ||
            binding.type === 'repeat-sheet' ||
            binding.type === 'subtemplate',
        );
      if (hasAdvanced) {
        const advanced = expandAdvancedTemplate(
          request.template,
          request.data,
          limits,
          environment.formatters ?? {},
          request.signal,
        );
        expansion = {
          document: advanced.document,
          diagnostics: advanced.diagnostics,
          insertedRows: new Map<string, readonly RowInsertion[]>(),
          repeatPageBreaks: advanced.forcedPageBreaks,
          structuralMappings: advanced.structuralMappings,
          objectMappings: advanced.objectMappings,
        };
      } else {
        const basic = expandTemplate(
          request.template.sourceDocument,
          request.template.ir.template,
          request.template.ir.bindings,
          request.data,
          environment.formatters ?? {},
          limits,
          request.signal,
        );
        expansion = { ...basic, structuralMappings: [], objectMappings: [] };
      }
    } catch (cause) {
      if (!(cause instanceof TemplateExpressionError)) throw cause;
      return failAfterResources(
        freeze({
          diagnostics: [
            renderDiagnostic(
              cause.code,
              cause.code === 'FORMATTER_FAILED'
                ? 'A template formatter failed during rendering'
                : cause.message,
            ),
          ],
        }),
      );
    }
    if (isAborted(request.signal)) return failAfterResources(abortResult());
    if (expansion.document === undefined)
      return failAfterResources(freeze({ diagnostics: expansion.diagnostics }));
    const expansionDiagnostics = expansion.diagnostics.map((diagnostic) =>
      request.missingValue === 'warning-and-blank' && diagnostic.code === 'MISSING_DATA'
        ? { ...diagnostic, severity: 'warning' as const }
        : diagnostic,
    );
    if (expansionDiagnostics.some(({ severity }) => severity === 'error')) {
      return failAfterResources(freeze({ diagnostics: expansionDiagnostics }));
    }
    const formulaEngine = createFormulaEngine();
    const formulaProgram = formulaEngine.compile(expansion.document);
    const calculation = formulaEngine.recalculate(formulaProgram, [], {
      locale: environment.locale,
      timeZone: environment.timeZone,
      dateSystem: environment.dateSystem,
      clock: { now: () => environment.clock.getTime() },
      tick: 0,
      functionRegistryVersion: 'builtin-1',
    });
    const activeFilterViews = (request.activeFilterViews ?? []).flatMap(({ sheetId, viewId }) => {
      const view = expansion.document?.workbook.sheets
        .find(({ id }) => id === sheetId)
        ?.filterViews.find(({ id }) => id === viewId);
      return view === undefined ? [] : [view];
    });
    const viewProjections = new Map(
      activeFilterViews.map((view) => [
        view.range.sheetId,
        {
          ...applyDocumentFilterView({
            document: expansion.document!,
            formulaValues: calculation.values,
            view,
            locale: environment.locale,
            limits: { maxRows: limits.maxExpandedRows },
          }),
          startRow: view.range.start.row + 1,
          endRow: view.range.end.row,
        },
      ]),
    );
    const resolvedTargets = targets(
      expansion.document,
      profile,
      expansion.insertedRows,
      expansion.structuralMappings,
    );
    const pageSize = paper(profile);
    const headingSize = profile.showHeadings ? 20 : 0;
    const pagination = paginateTemplateTargets({
      targets: paginationTargets(
        resolvedTargets,
        profile,
        expansion.insertedRows,
        expansion.structuralMappings,
        viewProjections,
      ),
      paper: pageSize,
      margins: {
        ...profile.page.margins,
        top: profile.page.margins.top + headingSize,
        left: profile.page.margins.left + headingSize,
      },
      scale: profile.page.scale,
      manualBreaks: profile.manualBreaks
        .flatMap((pageBreak) =>
          resolvedTargets
            .filter(
              ({ sheet }) =>
                sheet.id === pageBreak.sheetId ||
                expansion.structuralMappings.some(
                  ({ source, generated }) =>
                    source.sheetId === pageBreak.sheetId && generated.sheetId === sheet.id,
                ),
            )
            .flatMap((target) => {
              const beforeRow = mappedRow(
                pageBreak.beforeRow,
                expansion.insertedRows.get(pageBreak.sheetId) ?? [],
              );
              const projected = projectedBreakIndex(
                target,
                beforeRow,
                viewProjections.get(target.sheet.id),
              );
              return projected === undefined ? [] : [{ targetId: target.id, beforeRow: projected }];
            }),
        )
        .concat(
          resolvedTargets.flatMap((target) =>
            (expansion.repeatPageBreaks.get(target.sheet.id) ?? [])
              .filter(
                (beforeRow) =>
                  beforeRow >= target.range.start.row && beforeRow <= target.range.end.row,
              )
              .flatMap((beforeRow) => {
                const projected = projectedBreakIndex(
                  target,
                  beforeRow,
                  viewProjections.get(target.sheet.id),
                );
                return projected === undefined
                  ? []
                  : [{ targetId: target.id, beforeRow: projected }];
              }),
          ),
        ),
      maxPages: limits.maxPages,
      signal: request.signal,
      deadline: start + limits.maxLayoutTimeMs,
    });
    if (Date.now() - start > limits.maxLayoutTimeMs || isAborted(request.signal)) {
      return failAfterResources(
        isAborted(request.signal)
          ? abortResult()
          : freeze({
              diagnostics: [
                renderDiagnostic('LAYOUT_TIME_EXCEEDED', 'Template layout exceeded its time limit'),
              ],
            }),
      );
    }
    if (pagination.diagnostics.some(({ severity }) => severity === 'error')) {
      return failAfterResources(
        freeze({ diagnostics: [...expansionDiagnostics, ...pagination.diagnostics] }),
      );
    }
    const cache = createPresentationCache({
      maximumEntries: Math.max(1, limits.maxExpandedCells),
      maximumBytes: Math.max(1_024, limits.maxExpandedCells * 1_024),
    });
    const presentation = createPresentationResolver({
      document: expansion.document,
      formulaProgram,
      formulaValues: calculation.values,
      cache,
      revisions: {
        document: 1,
        calculation: 1,
        condition: 1,
        style: 1,
        environment: 1,
        view: activeFilterViews.length,
      },
      environment: {
        locale: environment.locale,
        timeZone: environment.timeZone,
        dateSystem: environment.dateSystem,
        target: 'print',
      },
      activeFilterViews,
    });
    const pageInputs = displayPages(
      pagination.pages,
      resolvedTargets,
      profile,
      (sheetId, row, column) => presentation.resolve({ sheetId, row, column }),
      expansion.insertedRows,
      expansion.structuralMappings,
      resourceBindings,
      resources,
      viewProjections,
      request.data,
      environment.clock,
      createPersistedVisualizationValueSource(
        expansion.document,
        `${request.currentDocumentHash}:print`,
        calculation.values,
      ),
    );
    const displayList = createPrintDisplayList({
      pages: pageInputs,
      fontMetrics: environment.fontMetrics,
      signal: request.signal,
      deadline: start + limits.maxLayoutTimeMs,
    });
    const diagnostics = [
      ...request.template.diagnostics,
      ...expansionDiagnostics,
      ...pagination.diagnostics,
      ...displayList.diagnostics,
    ];
    if (diagnostics.some(({ severity }) => severity === 'error')) {
      return failAfterResources(freeze({ diagnostics }));
    }
    const calculatedCells = expansion.document.workbook.sheets.flatMap((sheet) =>
      sheet.cells.map(({ row, column }) => {
        const address = { sheetId: sheet.id, row, column };
        return freeze({ address, value: presentation.resolve(address).value });
      }),
    );
    const document: GeneratedDocument = freeze({
      workbook: expansion.document.workbook,
      calculatedCells,
      worksheets: expansion.document.workbook.sheets.map((sheet) =>
        freeze({
          sheetId: sheet.id,
          visibility: sheet.visibility ?? 'visible',
          conditionalFormatting: freeze([...(sheet.conditionalFormatting ?? [])]),
        }),
      ),
      print: {
        pages: pagination.pages.map((page) => ({
          id: page.id,
          index: page.index,
          targetId: page.targetId,
          width: page.width,
          height: page.height,
          rowStart: page.rowStart,
          rowEnd: page.rowEnd,
          columnStart: page.columnStart,
          columnEnd: page.columnEnd,
        })),
        displayList,
        profile,
      },
      resources,
      objects: generatedObjectMappings(expansion.document, expansion.objectMappings),
      diagnostics,
      metadata: {
        templateId: request.template.templateId,
        profileId: profile.id,
        sourceDocumentHash: request.template.sourceDocumentHash,
        locale: environment.locale,
        timeZone: environment.timeZone,
        generatedAt: environment.clock.toISOString(),
      },
    });
    transferred = true;
    return freeze({ document, diagnostics });
  } finally {
    if (!transferred) await resources.dispose();
  }
}
