import type { DocumentCellAddress, ResourceId } from '../document';
import type { DisplayRect } from '../print';

export interface ObjectOffset {
  readonly x: number;
  readonly y: number;
}

export type ObjectAnchor =
  | {
      readonly type: 'absolute';
      readonly rect: DisplayRect;
    }
  | {
      readonly type: 'one-cell';
      readonly cell: DocumentCellAddress;
      readonly offset: ObjectOffset;
      readonly size: { readonly width: number; readonly height: number };
    }
  | {
      readonly type: 'two-cell';
      readonly from: DocumentCellAddress & { readonly offset: ObjectOffset };
      readonly to: DocumentCellAddress & { readonly offset: ObjectOffset };
    };

interface ObjectBase {
  readonly id: string;
  readonly anchor: ObjectAnchor;
  readonly zIndex: number;
  readonly locked: boolean;
  readonly templateRepeat: 'shared' | 'per-item' | 'forbidden';
  readonly accessibility: { readonly name: string; readonly description?: string };
}

export type SheetObject =
  | (ObjectBase & {
      readonly kind: 'image';
      readonly resourceId: ResourceId;
      readonly fit?: 'contain' | 'cover' | 'fill';
    })
  | (ObjectBase & {
      readonly kind: 'text-box';
      readonly text: string;
      readonly style: {
        readonly color: string;
        readonly fontFamily: string;
        readonly fontSize: number;
        readonly horizontalAlign?: 'left' | 'center' | 'right';
      };
    });

export interface ObjectGeometry {
  rowOffset(row: number): number;
  columnOffset(column: number): number;
}

export type ObjectCoordinateTransform = {
  readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column';
  readonly sheetId: string;
  readonly index: number;
  readonly count: number;
};
