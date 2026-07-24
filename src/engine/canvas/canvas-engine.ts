import type { CellPoint, CellRange } from '../../core/types/coordinates';
import type { CellStyle, SheetData } from '../../core/types/workbook';
import { createLegacyPresentationResolver, type CellPresentation } from '../../presentation';
import { frozenQuadrants } from '../geometry/frozen-pane-geometry';
import type { ViewportMetrics } from '../ports';
import { configuredCellDefaultStyle, paintCells, paintFilterOverlays } from './cell-painter';
import { currentDevicePixelRatio, DrawContext } from './draw-context';
import type { CanvasSurfacePort, TextMeasurementPort } from './draw-context';
import { createSparseCellScanBudget, paintGrid, paneCells, paneGridIndexes } from './grid-painter';
import { paintHeaders } from './header-painter';
import { RenderScheduler } from './render-scheduler';
import type { AnimationFramePort } from './render-scheduler';
import { paintCommands, paintObjects } from './object-painter';
import type { PrintDisplayCommand } from '../../print';
import { paintSelection } from './selection-painter';
import { paintTemplateDecorations } from './template-decoration-painter';
import type { ScreenObjectProjection } from '../../objects';
import type { FrozenQuadrantKind } from '../geometry/frozen-pane-geometry';

export interface TemplateCanvasDecoration {
  readonly range: CellRange;
  readonly kind: 'value' | 'repeat' | 'print';
  readonly label: string;
  readonly invalid?: boolean;
}

export type { CanvasSurfacePort, TextMeasurementPort } from './draw-context';
export type { AnimationFramePort } from './render-scheduler';

export interface CanvasRenderSnapshot {
  readonly sheet: Readonly<SheetData>;
  readonly viewport: ViewportMetrics;
  readonly selection?: CellRange;
  readonly invalidCells?: readonly CellPoint[];
  readonly templateDecorations?: readonly TemplateCanvasDecoration[];
  /** Revision-bound persisted chart and sparkline commands in worksheet coordinates. */
  readonly visualizationCommands?: readonly PrintDisplayCommand[];
  readonly showGrid?: boolean;
  /** Visible persistent objects already projected through shared display geometry. */
  readonly objects?: readonly ScreenObjectProjection[];
  /** Objects independently projected for each frozen-pane coordinate space. */
  readonly objectPanes?: readonly {
    readonly kind: FrozenQuadrantKind;
    readonly objects: readonly ScreenObjectProjection[];
  }[];
  /** Transient object selection; never serialized into the document. */
  readonly selectedObjectId?: string;
  /** Shared presentation batch for the visible document revision. */
  readonly presentations?: {
    readonly resolve: (point: CellPoint) => CellPresentation;
  };
}

export interface CanvasEngineOptions {
  readonly animationFrame?: AnimationFramePort;
  readonly defaultStyle?: CellStyle;
  readonly devicePixelRatio?: number;
  readonly measurement?: TextMeasurementPort;
  readonly onRenderError?: (cause: unknown) => void;
}

function canvasMeasurement(canvas: CanvasSurfacePort): TextMeasurementPort {
  return {
    measureText(text, font) {
      const context = canvas.getContext('2d');
      if (context === null) return text.length * 7;
      context.save();
      context.font = font;
      const width = context.measureText(text).width;
      context.restore();
      return width;
    },
  };
}

export class CanvasEngine {
  private readonly draw: DrawContext;
  private readonly defaultStyle: CellStyle;
  private readonly onRenderError: ((cause: unknown) => void) | undefined;
  private readonly scheduler: RenderScheduler;
  private latest: CanvasRenderSnapshot | null = null;
  private disposed = false;

  constructor(canvas: CanvasSurfacePort, options: Readonly<CanvasEngineOptions> = {}) {
    this.defaultStyle = configuredCellDefaultStyle(options.defaultStyle);
    this.onRenderError = options.onRenderError;
    this.draw = new DrawContext(
      canvas,
      options.devicePixelRatio ?? currentDevicePixelRatio(),
      options.measurement ?? canvasMeasurement(canvas),
    );
    this.scheduler = new RenderScheduler(options.animationFrame);
  }

  render(snapshot: CanvasRenderSnapshot): void {
    if (this.disposed) return;
    this.latest = snapshot;
    this.scheduler.schedule(() => this.paintScheduled());
  }

  private paintScheduled(): void {
    try {
      this.paintLatest();
    } catch (cause) {
      this.latest = null;
      if (this.onRenderError === undefined) throw cause;
      this.onRenderError(cause);
    }
  }

  private paintLatest(): void {
    const snapshot = this.latest;
    if (snapshot === null || this.disposed) return;
    const { viewport } = snapshot;
    const visibleRows = new Set<number>();
    const visibleColumns = new Set<number>();
    const scanBudget = createSparseCellScanBudget();
    const plans = frozenQuadrants(viewport.freeze, viewport).map((pane) => {
      const indexes = paneGridIndexes(pane, viewport);
      const cells = paneCells(viewport, indexes, snapshot.sheet, scanBudget);
      for (const row of indexes.rows) visibleRows.add(row);
      for (const column of indexes.columns) visibleColumns.add(column);
      return { pane, indexes, cells };
    });
    this.draw.resize(viewport.width, viewport.height);
    this.draw.clear(viewport.width, viewport.height);
    const legacyPresentations = createLegacyPresentationResolver(snapshot.sheet, this.defaultStyle);
    const presentations =
      snapshot.presentations === undefined
        ? legacyPresentations
        : {
            resolve: (point: CellPoint) => snapshot.presentations!.resolve(point),
          };
    for (const { pane, indexes, cells } of plans) {
      this.draw.withClip(
        pane,
        () => {
          if (snapshot.showGrid !== false) paintGrid(this.draw, indexes, viewport);
          paintCells(this.draw, snapshot, cells, presentations);
          paintFilterOverlays(this.draw, snapshot, indexes.rows, indexes.columns);
          paintCommands(this.draw.context, snapshot.visualizationCommands ?? []);
          paintSelection(this.draw, snapshot.selection, viewport, pane.kind);
          paintTemplateDecorations(
            this.draw,
            snapshot.templateDecorations ?? [],
            viewport,
            pane.kind,
          );
        },
        {
          x:
            viewport.rowHeaderWidth -
            (pane.kind === 'top' || pane.kind === 'body' ? viewport.scroll.x : 0),
          y:
            viewport.columnHeaderHeight -
            (pane.kind === 'left' || pane.kind === 'body' ? viewport.scroll.y : 0),
        },
      );
      const paneObjects =
        snapshot.objectPanes?.find(({ kind }) => kind === pane.kind)?.objects ??
        snapshot.objects ??
        [];
      if (paneObjects.length > 0) {
        this.draw.withClip(pane, () => {
          paintObjects(this.draw, paneObjects, snapshot.selectedObjectId);
        });
      }
    }
    paintHeaders(
      this.draw,
      viewport,
      Array.from(visibleRows),
      Array.from(visibleColumns),
      snapshot.selection,
      snapshot.sheet,
    );
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.latest = null;
    this.scheduler.dispose();
  }
}
