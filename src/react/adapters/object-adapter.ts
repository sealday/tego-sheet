import type { ResourceMetadata } from '../../document';
import { frozenQuadrants, type FrozenQuadrantKind, type ViewportMetrics } from '../../engine';
import {
  projectObjectsToScreen,
  type ResolvedScreenResource,
  type ScreenObjectProjection,
  type SheetObject,
} from '../../objects';

/** Screen object projections scoped to one frozen-pane clip. */
export interface ScreenObjectPane {
  readonly kind: FrozenQuadrantKind;
  readonly objects: readonly ScreenObjectProjection[];
}

/** Projects document objects into current Canvas coordinates without resolving external resources. */
export function projectSheetObjectsToViewport(
  objects: readonly SheetObject[],
  resources: readonly ResourceMetadata[],
  viewport: ViewportMetrics,
  resolvedResources?: Readonly<Record<string, ResolvedScreenResource>>,
): readonly ScreenObjectProjection[] {
  const visible = new Map<string, ScreenObjectProjection>();
  for (const pane of projectSheetObjectPanesToViewport(
    objects,
    resources,
    viewport,
    resolvedResources,
  )) {
    for (const projection of pane.objects) {
      const previous = visible.get(projection.object.id);
      visible.set(
        projection.object.id,
        previous === undefined
          ? projection
          : {
              ...previous,
              bounds: unionRect(previous.bounds, projection.bounds),
            },
      );
    }
  }
  return [...visible.values()].sort(
    (left, right) =>
      left.object.zIndex - right.object.zIndex || compareCodeUnits(left.object.id, right.object.id),
  );
}

/** Projects each object through one consistent coordinate transform per frozen pane. */
export function projectSheetObjectPanesToViewport(
  objects: readonly SheetObject[],
  resources: readonly ResourceMetadata[],
  viewport: ViewportMetrics,
  resolvedResources?: Readonly<Record<string, ResolvedScreenResource>>,
): readonly ScreenObjectPane[] {
  const persistentObjects = new Map(objects.map((object) => [object.id, object]));
  return frozenQuadrants(viewport.freeze, viewport).map((pane) => {
    const scrollX = pane.kind === 'top' || pane.kind === 'body' ? viewport.scroll.x : 0;
    const scrollY = pane.kind === 'left' || pane.kind === 'body' ? viewport.scroll.y : 0;
    const screenObjects = objects
      .filter((object) => object.anchor.type !== 'absolute' || pane.kind === 'body')
      .map((object): SheetObject => {
        if (object.anchor.type !== 'absolute') return object;
        return {
          ...object,
          anchor: {
            type: 'absolute',
            rect: {
              ...object.anchor.rect,
              x: viewport.rowHeaderWidth + object.anchor.rect.x - scrollX,
              y: viewport.columnHeaderHeight + object.anchor.rect.y - scrollY,
            },
          },
        };
      });
    const projected = projectObjectsToScreen(screenObjects, {
      resources,
      ...(resolvedResources === undefined ? {} : { resolvedResources }),
      geometry: {
        columnOffset(column) {
          return viewport.rowHeaderWidth + viewport.model.columnOffset(column) - scrollX;
        },
        rowOffset(row) {
          const visualRow = viewport.model.visualIndexOfRow(row);
          return viewport.columnHeaderHeight + viewport.model.rowOffset(visualRow) - scrollY;
        },
      },
      viewport: {
        x: pane.left,
        y: pane.top,
        width: pane.width,
        height: pane.height,
      },
    }).map((projection) => ({
      ...projection,
      object: persistentObjects.get(projection.object.id) ?? projection.object,
    }));
    return { kind: pane.kind, objects: projected };
  });
}

function unionRect(
  left: ScreenObjectProjection['bounds'],
  right: ScreenObjectProjection['bounds'],
): ScreenObjectProjection['bounds'] {
  const x = Math.min(left.x, right.x);
  const y = Math.min(left.y, right.y);
  const rightEdge = Math.max(left.x + left.width, right.x + right.width);
  const bottomEdge = Math.max(left.y + left.height, right.y + right.height);
  return { x, y, width: rightEdge - x, height: bottomEdge - y };
}

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}
