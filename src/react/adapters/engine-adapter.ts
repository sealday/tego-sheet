import {
  assertClipboardResourceLimit,
  parseA1,
  type LocaleDefinition,
  type Selection,
  type SheetId,
  type SheetOptions,
} from '../../core';
import type { SpreadsheetControllerSnapshot } from '../../core/controller/spreadsheet-document-controller';
import type { DocumentSheetId, FilterView } from '../../document';
import {
  CanvasEngine,
  clampScroll,
  createSelectionState,
  createSheetGridModel,
  createViewportMetrics,
  moveSelection,
  normalizeSelection,
  overlayAnchor,
  type OverlayAnchor,
  type SelectionDirection,
  type CanvasRenderSnapshot,
  type InteractionSnapshot,
  type ScrollState,
  type SelectionState,
  type TemplateCanvasDecoration,
  type ViewportMetrics,
} from '../../engine';
import { createPresentationCache, createPresentationResolver } from '../../presentation';
import { createPresentationValidationResolver } from './presentation-adapter';
import { projectSheetObjectsToViewport } from './object-adapter';
import { applyDocumentFilterView } from '../../views';

export interface EngineAdapterOptions {
  readonly root: HTMLElement;
  readonly canvas: HTMLCanvasElement;
  readonly onRenderError?: (cause: unknown) => void;
  readonly sheetOptions?: SheetOptions;
  readonly showGrid?: boolean;
  readonly locale?: LocaleDefinition;
  readonly getActiveFilterView?: (sheet: SheetId) => FilterView | undefined;
  readonly getFilterViewRevision?: () => number;
}

export interface EngineAdapter {
  readonly interactionSnapshot: () => InteractionSnapshot | null;
  readonly overlayAnchor: (
    point: Readonly<{ readonly row: number; readonly column: number }>,
  ) => OverlayAnchor | null;
  readonly ensureVisible: (
    point: Readonly<{ readonly row: number; readonly column: number }>,
  ) => OverlayAnchor | null;
  readonly nextSelection: (direction: SelectionDirection) => Readonly<{
    readonly state: SelectionState;
    readonly selection: Selection;
  }> | null;
  readonly publicSelection: () => Selection | null;
  readonly readSelection: (selection: Selection) => readonly (readonly string[])[];
  readonly refresh: (snapshot: SpreadsheetControllerSnapshot) => void;
  readonly render: (snapshot: SpreadsheetControllerSnapshot, activeSheet: SheetId | null) => void;
  readonly recalculateLayout: () => void;
  readonly setScroll: (scroll: ScrollState) => void;
  readonly setSelection: (selection: SelectionState) => void;
  readonly setSelectedObject: (objectId: string | null) => void;
  readonly stageSelection: (selection: SelectionState) => Selection | null;
  readonly updateReadOnly: (readOnly: boolean) => void;
  readonly updateLiveOptions: (options: Readonly<{ readonly showGrid?: boolean }>) => void;
  readonly updateTemplateDecorations: (decorations: readonly TemplateCanvasDecoration[]) => void;
  readonly dispose: () => void;
}

function dimensions(
  surface: HTMLElement,
  fallback: HTMLElement,
): { readonly width: number; readonly height: number } {
  const rect =
    typeof surface.getBoundingClientRect === 'function'
      ? surface.getBoundingClientRect()
      : undefined;
  const fallbackRect = fallback.getBoundingClientRect();
  return {
    width: Math.max(
      0,
      surface.clientWidth || rect?.width || fallback.clientWidth || fallbackRect.width || 0,
    ),
    height: Math.max(
      0,
      surface.clientHeight || rect?.height || fallback.clientHeight || fallbackRect.height || 0,
    ),
  };
}

function clippedFreeze(value: string | undefined): {
  readonly row: number;
  readonly column: number;
} {
  return parseA1(value ?? 'A1');
}

export function createEngineAdapter(options: EngineAdapterOptions): EngineAdapter {
  let latestSnapshot: SpreadsheetControllerSnapshot | null = null;
  let failedSnapshot: SpreadsheetControllerSnapshot | null = null;
  const engine = new CanvasEngine(options.canvas, {
    defaultStyle: options.sheetOptions?.defaultStyle,
    ...(options.onRenderError === undefined
      ? {}
      : {
          onRenderError: (cause: unknown) => {
            if (latestSnapshot === null || failedSnapshot === latestSnapshot) return;
            failedSnapshot = latestSnapshot;
            options.onRenderError?.(cause);
          },
        }),
  });
  let activeSheet: SheetId | null = null;
  let viewport: ViewportMetrics | null = null;
  let selection: SelectionState | null = null;
  let selectedObjectId: string | null = null;
  let disposed = false;
  let liveReadOnly: boolean | null = null;
  let showGrid = options.showGrid;
  let templateDecorations: readonly TemplateCanvasDecoration[] = [];
  const presentationCache = createPresentationCache({
    maximumEntries: 10_000,
    maximumBytes: 8 * 1024 * 1024,
  });

  const activeIndex = (): number => {
    if (latestSnapshot === null || activeSheet === null) return -1;
    return latestSnapshot.sheets.findIndex((sheet) => sheet.id === activeSheet);
  };

  const paint = () => {
    if (
      disposed ||
      latestSnapshot === null ||
      failedSnapshot === latestSnapshot ||
      viewport === null
    )
      return;
    const index = activeIndex();
    const sheet = index < 0 ? undefined : latestSnapshot.projection[index];
    const documentSheet = index < 0 ? undefined : latestSnapshot.document.workbook.sheets[index];
    if (sheet === undefined || documentSheet === undefined) return;
    const resolver = createPresentationResolver({
      document: latestSnapshot.document,
      formulaValues: new Map(
        latestSnapshot.calculation.values.map(({ address, value }) => [address, value]),
      ),
      formulaSpillAnchors: new Map(
        latestSnapshot.calculation.spillAnchors.map(({ address, anchor }) => [address, anchor]),
      ),
      cache: presentationCache,
      validation: createPresentationValidationResolver(latestSnapshot),
      revisions: {
        document: latestSnapshot.revision,
        calculation: latestSnapshot.calculation.revision,
        condition: latestSnapshot.revision,
        style: latestSnapshot.revision,
        environment: latestSnapshot.revision,
        view: options.getFilterViewRevision?.() ?? 0,
      },
      environment: {
        locale:
          options.locale?.id ?? latestSnapshot.document.workbook.settings.localeHint ?? 'en-US',
        timeZone: 'UTC',
        dateSystem: latestSnapshot.document.workbook.settings.dateSystem,
        target: 'screen',
      },
      ...(activeSheet === null || options.getActiveFilterView === undefined
        ? {}
        : { activeFilterView: options.getActiveFilterView(activeSheet) }),
    });
    const renderSnapshot: CanvasRenderSnapshot = {
      sheet,
      viewport,
      objects: projectSheetObjectsToViewport(
        documentSheet.objects,
        latestSnapshot.document.resources.items,
        viewport,
      ),
      presentations: {
        resolve: ({ row, column }) =>
          resolver.resolve({
            sheetId: documentSheet.id as DocumentSheetId,
            row,
            column,
          }),
      },
      ...(selection === null ? {} : { selection: selection.range }),
      ...(selectedObjectId === null ? {} : { selectedObjectId }),
      showGrid,
      templateDecorations,
    };
    engine.render(renderSnapshot);
  };

  const rebuild = (paintNow = true) => {
    if (disposed || latestSnapshot === null) return;
    const index = activeIndex();
    const sheet = index < 0 ? undefined : latestSnapshot.projection[index];
    const documentSheet = index < 0 ? undefined : latestSnapshot.document.workbook.sheets[index];
    if (sheet === undefined || documentSheet === undefined) {
      viewport = null;
      selection = null;
      return;
    }
    const activeView =
      activeSheet === null ? undefined : options.getActiveFilterView?.(activeSheet);
    const derivedRows =
      activeView === undefined
        ? undefined
        : applyDocumentFilterView({
            document: latestSnapshot.document,
            formulaValues: new Map(
              latestSnapshot.calculation.values.map(({ address, value }) => [address, value]),
            ),
            view: activeView,
            locale:
              options.locale?.id ?? latestSnapshot.document.workbook.settings.localeHint ?? 'en-US',
            limits: {
              maxRows: Math.max(1, activeView.range.end.row - activeView.range.start.row + 1),
            },
          });
    const model = createSheetGridModel(sheet, {
      defaultRowHeight: options.sheetOptions?.rows?.defaultHeight,
      defaultColumnWidth: options.sheetOptions?.columns?.defaultWidth,
      locale: options.locale,
      ...(derivedRows === undefined
        ? {}
        : {
            derivedRows: {
              start: activeView!.range.start.row + 1,
              end: activeView!.range.end.row,
              ...derivedRows,
            },
          }),
    });
    const previousScroll = viewport?.scroll ?? { x: 0, y: 0 };
    viewport = createViewportMetrics(model, {
      ...dimensions(options.canvas, options.root),
      rowHeaderWidth: options.sheetOptions?.rowHeaderWidth,
      scroll: previousScroll,
      freeze: clippedFreeze(sheet.freeze),
    });
    const clippedScroll = clampScroll(previousScroll, viewport);
    if (clippedScroll.x !== previousScroll.x || clippedScroll.y !== previousScroll.y) {
      viewport = createViewportMetrics(model, { ...viewport, scroll: clippedScroll });
    }
    if (model.rowCount === 0 || model.columnCount === 0) {
      selection = null;
    } else {
      selection = normalizeSelection(
        selection ?? createSelectionState({ row: 0, column: 0 }),
        model,
      );
    }
    if (paintNow) paint();
  };

  const ensureVisible = (
    point: Readonly<{ readonly row: number; readonly column: number }>,
    paintNow = true,
  ) => {
    if (
      disposed ||
      viewport === null ||
      point.row < 0 ||
      point.column < 0 ||
      point.row >= viewport.model.rowCount ||
      point.column >= viewport.model.columnCount
    )
      return null;
    const model = viewport.model;
    const frozenWidth = model.columnOffset(viewport.freeze.column);
    const frozenHeight = model.rowOffset(viewport.freeze.row);
    const bodyWidth = Math.max(0, viewport.width - viewport.rowHeaderWidth - frozenWidth);
    const bodyHeight = Math.max(0, viewport.height - viewport.columnHeaderHeight - frozenHeight);
    let x = viewport.scroll.x;
    let y = viewport.scroll.y;
    if (point.column >= viewport.freeze.column) {
      const start = model.columnOffset(point.column);
      const end = model.columnOffset(point.column + 1);
      const visibleStart = frozenWidth + x;
      const visibleEnd = visibleStart + bodyWidth;
      if (start < visibleStart) x = start - frozenWidth;
      else if (end > visibleEnd) x = end - frozenWidth - bodyWidth;
    }
    const visualRow = model.visualIndexOfRow(point.row);
    if (visualRow >= viewport.freeze.row) {
      const start = model.rowOffset(visualRow);
      const end = model.rowOffset(visualRow + 1);
      const visibleStart = frozenHeight + y;
      const visibleEnd = visibleStart + bodyHeight;
      if (start < visibleStart) y = start - frozenHeight;
      else if (end > visibleEnd) y = end - frozenHeight - bodyHeight;
    }
    const next = clampScroll({ x, y }, viewport);
    if (next.x !== viewport.scroll.x || next.y !== viewport.scroll.y) {
      viewport = createViewportMetrics(model, { ...viewport, scroll: next });
      if (paintNow) paint();
    }
    const range = { start: point, end: point };
    return overlayAnchor(range, viewport);
  };

  return {
    interactionSnapshot() {
      if (
        disposed ||
        latestSnapshot === null ||
        activeSheet === null ||
        viewport === null ||
        selection === null
      )
        return null;
      return {
        viewport,
        selection,
        sheet: activeSheet,
        readOnly: liveReadOnly ?? latestSnapshot.readOnly,
        epoch: latestSnapshot,
      };
    },
    overlayAnchor(point) {
      if (disposed || viewport === null) return null;
      const range = { start: point, end: point };
      return overlayAnchor(range, viewport);
    },
    ensureVisible(point) {
      return ensureVisible(point);
    },
    nextSelection(direction) {
      if (disposed || viewport === null || selection === null || activeSheet === null) return null;
      const next = moveSelection(selection, direction, viewport.model);
      return {
        state: next,
        selection: { sheet: activeSheet, range: next.range, active: next.active },
      };
    },
    publicSelection() {
      if (activeSheet === null || selection === null) return null;
      return {
        sheet: activeSheet,
        range: selection.range,
        active: selection.active,
      };
    },
    readSelection(requested) {
      if (latestSnapshot === null) return [];
      const index = latestSnapshot.sheets.findIndex((sheet) => sheet.id === requested.sheet);
      const sheet = index < 0 ? undefined : latestSnapshot.projection[index];
      if (sheet === undefined) return [];
      assertClipboardResourceLimit(requested.range);
      const output: string[][] = [];
      for (let row = requested.range.start.row; row <= requested.range.end.row; row += 1) {
        const values: string[] = [];
        for (
          let column = requested.range.start.column;
          column <= requested.range.end.column;
          column += 1
        ) {
          const rowData = sheet.rows?.[String(row)];
          const rowRecord =
            typeof rowData === 'object' && rowData !== null && !Array.isArray(rowData)
              ? (rowData as Readonly<Record<string, unknown>>)
              : undefined;
          const cells = rowRecord?.cells;
          const cell =
            typeof cells === 'object' && cells !== null && !Array.isArray(cells)
              ? (cells as Readonly<Record<string, unknown>>)[String(column)]
              : undefined;
          const cellRecord =
            typeof cell === 'object' && cell !== null && !Array.isArray(cell)
              ? (cell as Readonly<Record<string, unknown>>)
              : undefined;
          values.push(typeof cellRecord?.text === 'string' ? cellRecord.text : '');
        }
        output.push(values);
      }
      return output;
    },
    refresh(snapshot) {
      if (disposed) return;
      latestSnapshot = snapshot;
      rebuild(false);
    },
    render(snapshot, sheet) {
      if (disposed) return;
      latestSnapshot = snapshot;
      activeSheet = sheet;
      rebuild();
    },
    recalculateLayout() {
      failedSnapshot = null;
      rebuild();
    },
    setScroll(scroll) {
      if (disposed || viewport === null) return;
      viewport = createViewportMetrics(viewport.model, { ...viewport, scroll });
      paint();
    },
    setSelection(next) {
      if (disposed || viewport === null) return;
      selection = normalizeSelection(next, viewport.model);
      paint();
    },
    setSelectedObject(objectId) {
      if (disposed || selectedObjectId === objectId) return;
      selectedObjectId = objectId;
      paint();
    },
    stageSelection(next) {
      if (disposed || viewport === null || activeSheet === null) return null;
      selection = normalizeSelection(next, viewport.model);
      ensureVisible(selection.active, false);
      return {
        sheet: activeSheet,
        range: selection.range,
        active: selection.active,
      };
    },
    updateReadOnly(readOnly) {
      if (disposed) return;
      liveReadOnly = readOnly;
    },
    updateLiveOptions(next) {
      if (disposed || showGrid === next.showGrid) return;
      showGrid = next.showGrid;
      paint();
    },
    updateTemplateDecorations(next) {
      if (disposed) return;
      templateDecorations = next;
      paint();
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      latestSnapshot = null;
      failedSnapshot = null;
      viewport = null;
      selection = null;
      selectedObjectId = null;
      presentationCache.clear();
      engine.dispose();
    },
  };
}
