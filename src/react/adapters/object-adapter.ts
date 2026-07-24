import type { ResourceMetadata } from '../../document';
import type { ViewportMetrics } from '../../engine';
import {
  projectObjectsToScreen,
  type ScreenObjectProjection,
  type SheetObject,
} from '../../objects';

/** Projects document objects into current Canvas coordinates without resolving external resources. */
export function projectSheetObjectsToViewport(
  objects: readonly SheetObject[],
  resources: readonly ResourceMetadata[],
  viewport: ViewportMetrics,
): readonly ScreenObjectProjection[] {
  const persistentObjects = new Map(objects.map((object) => [object.id, object]));
  const screenObjects = objects.map((object): SheetObject => {
    if (object.anchor.type !== 'absolute') return object;
    return {
      ...object,
      anchor: {
        type: 'absolute',
        rect: {
          ...object.anchor.rect,
          x: viewport.rowHeaderWidth + object.anchor.rect.x - viewport.scroll.x,
          y: viewport.columnHeaderHeight + object.anchor.rect.y - viewport.scroll.y,
        },
      },
    };
  });
  return projectObjectsToScreen(screenObjects, {
    resources,
    geometry: {
      columnOffset(column) {
        return (
          viewport.rowHeaderWidth +
          viewport.model.columnOffset(column) -
          (column < viewport.freeze.column ? 0 : viewport.scroll.x)
        );
      },
      rowOffset(row) {
        const visualRow = viewport.model.visualIndexOfRow(row);
        return (
          viewport.columnHeaderHeight +
          viewport.model.rowOffset(visualRow) -
          (visualRow < viewport.freeze.row ? 0 : viewport.scroll.y)
        );
      },
    },
    viewport: {
      x: viewport.rowHeaderWidth,
      y: viewport.columnHeaderHeight,
      width: Math.max(0, viewport.width - viewport.rowHeaderWidth),
      height: Math.max(0, viewport.height - viewport.columnHeaderHeight),
    },
  }).map((projection) => ({
    ...projection,
    object: persistentObjects.get(projection.object.id) ?? projection.object,
  }));
}
