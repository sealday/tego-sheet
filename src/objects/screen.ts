import type { ResourceMetadata } from '../document/model/document';
import type { DisplayRect, PrintDisplayCommand } from '../print';
import { resolveObjectAnchor } from './anchors';
import { objectToDisplayCommands } from './display';
import type { ObjectGeometry, SheetObject } from './model';

/** Stable screen-rendering diagnostic attached to one projected object. */
export interface ObjectScreenDiagnostic {
  /** Machine-readable failure category. */
  readonly code: 'OBJECT_RESOURCE_MISSING';
  /** Object that could not use its resource. */
  readonly objectId: string;
  /** Resource that was absent or unsafe to load. */
  readonly resourceId: string;
  /** Human-readable failure detail. */
  readonly message: string;
}

/** Screen-ready object retaining its persistent model and shared display commands. */
export interface ScreenObjectProjection {
  /** Persistent object. Selection remains outside this value. */
  readonly object: SheetObject;
  /** Rotation-aware bounds used for viewport intersection and hit testing. */
  readonly bounds: DisplayRect;
  /** Renderer-neutral commands shared with print output. */
  readonly commands: readonly PrintDisplayCommand[];
  /** Recoverable rendering diagnostics. */
  readonly diagnostics: readonly ObjectScreenDiagnostic[];
  /** Decoded image representations referenced by this projection's commands. */
  readonly imageResources?: Readonly<
    Record<
      string,
      {
        readonly source: unknown;
        readonly width: number;
        readonly height: number;
      }
    >
  >;
}

/** Resource-pipeline result required before Canvas may render an image. */
export interface ResolvedScreenResource {
  /** Validated logical resource type. */
  readonly type: string;
  /** Validated canonical MIME type. */
  readonly mimeType: string;
  /** Host-decoded Canvas-compatible representation. */
  readonly decoded?: unknown;
  /** Validated decoded pixel width. */
  readonly width?: number;
  /** Validated decoded pixel height. */
  readonly height?: number;
}

/** Inputs required to project persistent objects into the visible screen layer. */
export interface ObjectScreenContext {
  /** Worksheet geometry used by all object output paths. */
  readonly geometry: ObjectGeometry;
  /** Persistent document resource metadata. */
  readonly resources: readonly ResourceMetadata[];
  /** Validated and decoded resources keyed by persistent resource ID. */
  readonly resolvedResources?: Readonly<Record<string, ResolvedScreenResource>>;
  /** Visible screen rectangle expressed in the same geometry coordinates. */
  readonly viewport: DisplayRect;
}

/** Projects visible objects using the same geometry and display commands as print. */
export function projectObjectsToScreen(
  objects: readonly SheetObject[],
  context: ObjectScreenContext,
): readonly ScreenObjectProjection[] {
  const resources = new Map(context.resources.map((resource) => [resource.id, resource]));
  return objects
    .map((object): ScreenObjectProjection | undefined => {
      const rect = resolveObjectAnchor(object.anchor, context.geometry);
      const bounds = rotatedBounds(rect, object.rotation ?? 0);
      if (!rectanglesIntersect(bounds, context.viewport)) return undefined;

      if (object.kind === 'image') {
        const resource = resources.get(object.resourceId);
        const resolved = context.resolvedResources?.[object.resourceId];
        if (
          resource === undefined ||
          !isSafeImageResource(resource) ||
          resolved?.type !== 'image' ||
          resolved.decoded === undefined ||
          resolved.width === undefined ||
          resolved.height === undefined
        ) {
          return {
            object,
            bounds,
            commands: rotateCommands(object, rect, unavailableImageCommands(rect)),
            diagnostics: [
              {
                code: 'OBJECT_RESOURCE_MISSING',
                objectId: object.id,
                resourceId: object.resourceId,
                message:
                  resource === undefined
                    ? `Object ${object.id} references missing resource ${object.resourceId}`
                    : !isSafeImageResource(resource)
                      ? `Object ${object.id} references an unsafe image resource ${object.resourceId}`
                      : `Object ${object.id} references unresolved image resource ${object.resourceId}`,
              },
            ],
          };
        }
      }

      return {
        object,
        bounds,
        commands: objectToDisplayCommands(object, {
          geometry: context.geometry,
          resources:
            object.kind === 'image'
              ? { [object.resourceId]: new Uint8Array(0) }
              : Object.create(null),
        }),
        diagnostics: [],
        ...(object.kind === 'image'
          ? {
              imageResources: {
                [object.resourceId]: {
                  source: context.resolvedResources![object.resourceId]!.decoded,
                  width: context.resolvedResources![object.resourceId]!.width!,
                  height: context.resolvedResources![object.resourceId]!.height!,
                },
              },
            }
          : {}),
      };
    })
    .filter((projection): projection is ScreenObjectProjection => projection !== undefined)
    .sort(
      (left, right) =>
        left.object.zIndex - right.object.zIndex ||
        compareCodeUnits(left.object.id, right.object.id),
    );
}

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function unavailableImageCommands(rect: DisplayRect): readonly PrintDisplayCommand[] {
  return [
    { kind: 'fill-rect', rect, color: '#f3f4f6' },
    { kind: 'stroke-rect', rect, color: '#9ca3af', width: 1 },
    {
      kind: 'text',
      text: 'Image unavailable',
      x: rect.x + rect.width / 2,
      y: rect.y + rect.height / 2,
      maxWidth: rect.width,
      fontFamily: 'sans-serif',
      fontSize: Math.min(12, Math.max(1, rect.height)),
      color: '#4b5563',
      horizontalAlign: 'center',
    },
  ];
}

function rotateCommands(
  object: SheetObject,
  rect: DisplayRect,
  commands: readonly PrintDisplayCommand[],
): readonly PrintDisplayCommand[] {
  if (object.rotation === undefined || object.rotation === 0) return commands;
  return [
    {
      kind: 'group',
      rotation: object.rotation,
      origin: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
      commands,
    },
  ];
}

function isSafeImageResource(resource: ResourceMetadata): boolean {
  if (resource.mimeType !== undefined && !resource.mimeType.startsWith('image/')) return false;
  if (resource.url === undefined) return true;
  const normalized = resource.url.trim().toLowerCase();
  if (normalized.startsWith('data:')) return normalized.startsWith('data:image/');
  return !/^(?:javascript|vbscript|file):/.test(normalized);
}

function rectanglesIntersect(left: DisplayRect, right: DisplayRect): boolean {
  return (
    left.x < right.x + right.width &&
    left.x + left.width > right.x &&
    left.y < right.y + right.height &&
    left.y + left.height > right.y
  );
}

function rotatedBounds(rect: DisplayRect, degrees: number): DisplayRect {
  if (degrees % 360 === 0) return rect;
  const radians = (degrees * Math.PI) / 180;
  const width =
    Math.abs(rect.width * Math.cos(radians)) + Math.abs(rect.height * Math.sin(radians));
  const height =
    Math.abs(rect.width * Math.sin(radians)) + Math.abs(rect.height * Math.cos(radians));
  return {
    x: cleanNumber(rect.x + (rect.width - width) / 2),
    y: cleanNumber(rect.y + (rect.height - height) / 2),
    width: cleanNumber(width),
    height: cleanNumber(height),
  };
}

function cleanNumber(value: number): number {
  const rounded = Math.round(value * 1_000_000_000_000) / 1_000_000_000_000;
  return Object.is(rounded, -0) ? 0 : rounded;
}
