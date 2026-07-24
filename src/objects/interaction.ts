import type { ObjectAnchor, SheetObject } from './model';

/** Keyboard-only object transform expressed in device-independent units. */
export type ObjectKeyboardTransform =
  | {
      readonly type: 'move' | 'resize';
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly type: 'rotate';
      readonly degrees: number;
    };

/** Produces an immutable command payload, or no payload when the selected object is locked. */
export function transformObjectByKeyboard(
  object: SheetObject,
  transform: ObjectKeyboardTransform,
): SheetObject | undefined {
  if (object.locked) return undefined;
  if (transform.type === 'rotate') {
    return {
      ...object,
      rotation: normalizeRotation((object.rotation ?? 0) + transform.degrees),
    };
  }
  return {
    ...object,
    anchor: transformAnchor(object.anchor, transform),
  };
}

function transformAnchor(
  anchor: ObjectAnchor,
  transform: Extract<ObjectKeyboardTransform, { readonly type: 'move' | 'resize' }>,
): ObjectAnchor {
  if (anchor.type === 'absolute') {
    return {
      ...anchor,
      rect:
        transform.type === 'move'
          ? { ...anchor.rect, x: anchor.rect.x + transform.x, y: anchor.rect.y + transform.y }
          : {
              ...anchor.rect,
              width: positiveSize(anchor.rect.width + transform.x),
              height: positiveSize(anchor.rect.height + transform.y),
            },
    };
  }
  if (anchor.type === 'one-cell') {
    return transform.type === 'move'
      ? {
          ...anchor,
          offset: { x: anchor.offset.x + transform.x, y: anchor.offset.y + transform.y },
        }
      : {
          ...anchor,
          size: {
            width: positiveSize(anchor.size.width + transform.x),
            height: positiveSize(anchor.size.height + transform.y),
          },
        };
  }
  if (transform.type === 'move') {
    return {
      ...anchor,
      from: {
        ...anchor.from,
        offset: {
          x: anchor.from.offset.x + transform.x,
          y: anchor.from.offset.y + transform.y,
        },
      },
      to: {
        ...anchor.to,
        offset: {
          x: anchor.to.offset.x + transform.x,
          y: anchor.to.offset.y + transform.y,
        },
      },
    };
  }
  return {
    ...anchor,
    to: {
      ...anchor.to,
      offset: {
        x: anchor.to.offset.x + transform.x,
        y: anchor.to.offset.y + transform.y,
      },
    },
  };
}

function positiveSize(value: number): number {
  return Math.max(1, value);
}

function normalizeRotation(value: number): number {
  return ((value % 360) + 360) % 360;
}
