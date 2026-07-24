import type { KeyboardEvent } from 'react';
import {
  transformObjectByKeyboard,
  type ObjectKeyboardTransform,
  type ScreenObjectProjection,
  type SheetObject,
} from '../../objects';

export interface AccessibilityObjectsProps {
  /** Visible objects in stable paint order. */
  readonly objects: readonly ScreenObjectProjection[];
  /** Transient selected object identifier. */
  readonly selectedObjectId: string | null;
  /** Selects an object without modifying its persistent data. */
  readonly onSelect: (objectId: string) => void;
  /** Dispatches one immutable object update through the document controller. */
  readonly onChange: (object: SheetObject) => void;
}

/** Keyboard-accessible semantic mirror for the Canvas object layer. */
export function AccessibilityObjects({
  objects,
  selectedObjectId,
  onSelect,
  onChange,
}: AccessibilityObjectsProps) {
  return (
    <div role="listbox" aria-label="Sheet objects">
      {objects.map(({ object }) => (
        <div
          key={object.id}
          role="option"
          tabIndex={0}
          aria-label={object.accessibility.name}
          aria-description={object.accessibility.description}
          aria-selected={selectedObjectId === object.id}
          aria-readonly={object.locked}
          data-object-id={object.id}
          data-locked={object.locked}
          onClick={() => onSelect(object.id)}
          onFocus={() => onSelect(object.id)}
          onKeyDown={(event) => {
            const transform = keyboardTransform(event);
            if (transform === undefined) return;
            event.preventDefault();
            onSelect(object.id);
            const next = transformObjectByKeyboard(object, transform);
            if (next !== undefined) onChange(next);
          }}
        />
      ))}
    </div>
  );
}

function keyboardTransform(event: KeyboardEvent): ObjectKeyboardTransform | undefined {
  const direction = arrowDirection(event.key);
  if (direction === undefined) return undefined;
  if (event.altKey) {
    return {
      type: 'rotate',
      degrees: direction.x === 0 ? direction.y : direction.x,
    };
  }
  return {
    type: event.shiftKey ? 'resize' : 'move',
    ...direction,
  };
}

function arrowDirection(key: string): { readonly x: number; readonly y: number } | undefined {
  if (key === 'ArrowLeft') return { x: -1, y: 0 };
  if (key === 'ArrowRight') return { x: 1, y: 0 };
  if (key === 'ArrowUp') return { x: 0, y: -1 };
  if (key === 'ArrowDown') return { x: 0, y: 1 };
  return undefined;
}
