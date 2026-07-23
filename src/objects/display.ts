import type { PrintImageCommand, PrintTextCommand } from '../print';
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
          readonly bytes: Uint8Array | readonly number[];
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
): readonly (PrintImageCommand | PrintTextCommand)[] {
  const rect = resolveObjectAnchor(object.anchor, context.geometry);
  if (object.kind === 'image') {
    if (context.resources[object.resourceId] === undefined) {
      throw new SheetObjectError(
        'OBJECT_RESOURCE_MISSING',
        `Object ${object.id} references missing resource ${object.resourceId}`,
      );
    }
    return [
      {
        kind: 'image',
        resourceId: object.resourceId,
        rect,
        fit: object.fit ?? 'contain',
      },
    ];
  }
  return [
    {
      kind: 'text',
      text: object.text,
      x: rect.x,
      y: rect.y,
      maxWidth: rect.width,
      fontFamily: object.style.fontFamily,
      fontSize: object.style.fontSize,
      color: object.style.color,
      horizontalAlign: object.style.horizontalAlign ?? 'left',
    },
  ];
}
