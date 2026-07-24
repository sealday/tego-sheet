import type { DisplayRect, PrintDisplayCommand } from '../print';
import type { ObjectGeometry, SheetObject } from './model';
import { resolveObjectAnchor } from './anchors';

/** Stable object-to-output conversion failure. */
export class SheetObjectError extends Error {
  /** Creates a stable object rendering error. */
  constructor(
    /** Machine-readable object failure category. */
    readonly code: 'OBJECT_RESOURCE_MISSING',
    message: string,
  ) {
    super(message);
    this.name = 'SheetObjectError';
  }
}

/** Explicit resources and geometry used for object rendering. */
export interface ObjectDisplayContext {
  /** Already-resolved document resources keyed by stable ID. */
  readonly resources: Readonly<
    Record<
      string,
      | Uint8Array
      | {
          /** Encoded resource bytes. */
          readonly bytes: Uint8Array | readonly number[];
          /** Optional media type associated with the encoded bytes. */
          readonly mimeType?: string;
        }
    >
  >;
  /** Worksheet geometry used to resolve anchors. */
  readonly geometry: ObjectGeometry;
}

/** Converts one object to safe renderer-neutral display commands. */
export function objectToDisplayCommands(
  object: SheetObject,
  context: ObjectDisplayContext,
): readonly PrintDisplayCommand[] {
  const rect = resolveObjectAnchor(object.anchor, context.geometry);
  const rotate = (commands: readonly PrintDisplayCommand[]): readonly PrintDisplayCommand[] =>
    object.rotation === undefined || object.rotation === 0
      ? commands
      : [
          {
            kind: 'group',
            rotation: object.rotation,
            origin: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
            commands,
          },
        ];
  if (object.kind === 'image') {
    if (context.resources[object.resourceId] === undefined) {
      throw new SheetObjectError(
        'OBJECT_RESOURCE_MISSING',
        `Object ${object.id} references missing resource ${object.resourceId}`,
      );
    }
    return rotate([
      {
        kind: 'image',
        resourceId: object.resourceId,
        rect,
        fit: object.fit ?? 'contain',
      },
    ]);
  }
  if (object.kind === 'shape') {
    const stroke = object.style.stroke ?? '#000000';
    const width = object.style.strokeWidth ?? 1;
    if (object.shape === 'line') {
      return rotate([
        {
          kind: 'line',
          x1: rect.x,
          y1: rect.y,
          x2: rect.x + rect.width,
          y2: rect.y + rect.height,
          color: stroke,
          width,
        },
      ]);
    }
    if (object.shape === 'ellipse') {
      return rotate([
        {
          kind: 'path',
          data: ellipsePath(rect),
          ...(object.style.fill === undefined ? {} : { fill: object.style.fill }),
          stroke,
          width,
        },
      ]);
    }
    const commands: PrintDisplayCommand[] = [];
    if (object.style.fill !== undefined) {
      commands.push({ kind: 'fill-rect', rect, color: object.style.fill });
    }
    commands.push({ kind: 'stroke-rect', rect, color: stroke, width });
    return rotate(commands);
  }
  const text: PrintDisplayCommand = {
    kind: 'text',
    text: object.text,
    x: rect.x,
    y: rect.y + Math.min(rect.height, object.style.fontSize),
    maxWidth: rect.width,
    fontFamily: object.style.fontFamily,
    fontSize: object.style.fontSize,
    color: object.style.color,
    horizontalAlign: object.style.horizontalAlign ?? 'left',
  };
  return rotate([
    {
      kind: 'clip',
      rect,
      commands: [text],
    },
  ]);
}

function ellipsePath(rect: DisplayRect): string {
  const radiusX = rect.width / 2;
  const radiusY = rect.height / 2;
  const centerX = rect.x + radiusX;
  const bottom = rect.y + rect.height;
  return `M ${centerX} ${rect.y} A ${radiusX} ${radiusY} 0 1 0 ${centerX} ${bottom} A ${radiusX} ${radiusY} 0 1 0 ${centerX} ${rect.y} Z`;
}
