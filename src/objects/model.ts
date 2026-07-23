import type { DocumentCellAddress, ResourceId } from '../document';
import type { DisplayRect } from '../print';

/** Device-independent offset from an anchor cell. */
export interface ObjectOffset {
  /** Horizontal offset. */
  readonly x: number;
  /** Vertical offset. */
  readonly y: number;
}

/** Logical positioning modes shared by all floating objects. */
export type ObjectAnchor =
  | {
      /** Uses page-like absolute geometry. */
      readonly type: 'absolute';
      /** Fixed device-independent rectangle. */
      readonly rect: DisplayRect;
    }
  | {
      /** Positions a fixed-size object relative to one cell. */
      readonly type: 'one-cell';
      /** Anchor cell. */
      readonly cell: DocumentCellAddress;
      /** Offset from the cell origin. */
      readonly offset: ObjectOffset;
      /** Fixed object size. */
      readonly size: {
        /** Fixed width. */
        readonly width: number;
        /** Fixed height. */
        readonly height: number;
      };
    }
  | {
      /** Sizes an object between two cell markers. */
      readonly type: 'two-cell';
      /** Top-left marker. */
      readonly from: DocumentCellAddress & {
        /** Offset from the top-left cell origin. */
        readonly offset: ObjectOffset;
      };
      /** Bottom-right marker. */
      readonly to: DocumentCellAddress & {
        /** Offset from the bottom-right cell origin. */
        readonly offset: ObjectOffset;
      };
    };

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
  readonly accessibility: { readonly name: string; readonly description?: string };
}

/** Supported persistent worksheet floating objects. */
export type SheetObject =
  | (ObjectBase & {
      /** Raster image object. */
      readonly kind: 'image';
      /** Document-owned image resource. */
      readonly resourceId: ResourceId;
      /** Image fitting policy. */
      readonly fit?: 'contain' | 'cover' | 'fill';
    })
  | (ObjectBase & {
      /** Plain-text box object. */
      readonly kind: 'text-box';
      /** Plain text that is never interpreted as markup. */
      readonly text: string;
      /** Deterministic text presentation. */
      readonly style: {
        /** CSS-compatible text color. */
        readonly color: string;
        /** Resolved font family. */
        readonly fontFamily: string;
        /** Font size in device-independent units. */
        readonly fontSize: number;
        /** Optional horizontal alignment. */
        readonly horizontalAlign?: 'left' | 'center' | 'right';
      };
    });

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
