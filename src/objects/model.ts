import type { ObjectAnchor, ObjectRect, SheetObject } from '../document/model/document';

export type { ObjectAnchor, ObjectRect, SheetObject };

/** Device-independent offset from an anchor cell. */
export type ObjectOffset = Pick<ObjectRect, 'x' | 'y'>;

/** Properties shared by every persistent floating object. */
export interface ObjectBase {
  /** Stable object identifier. */
  readonly id: string;
  /** Logical position and size. */
  readonly anchor: ObjectAnchor;
  /** Stable paint order within a worksheet. */
  readonly zIndex: number;
  /** Whether interactive editing is locked. */
  readonly locked: boolean;
  /** Object policy during template repetition. */
  readonly templateRepeat: 'shared' | 'per-item' | 'forbidden';
  /** Accessible object description. */
  readonly accessibility: {
    /** Accessible object name. */
    readonly name: string;
    /** Optional longer accessible description. */
    readonly description?: string;
  };
}

/** Sheet geometry provider used to resolve logical anchors. */
export interface ObjectGeometry {
  /** Returns the top edge of a row. */
  rowOffset(row: number): number;
  /** Returns the left edge of a column. */
  columnOffset(column: number): number;
}

/** Structural sheet operation applied to object anchors. */
export type ObjectCoordinateTransform = {
  /** Row or column structural operation. */
  readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column';
  /** Sheet receiving the operation. */
  readonly sheetId: string;
  /** First affected row or column. */
  readonly index: number;
  /** Number of inserted or deleted rows or columns. */
  readonly count: number;
};
