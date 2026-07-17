import type { JsonExtensible, JsonValue, SparseJsonCollection } from './json';
import type { ValidationOperator, ValidationType } from './validation';

/** Horizontal text alignment within a cell. */
export type HorizontalAlign = 'left' | 'center' | 'right';

/** Vertical text alignment within a cell. */
export type VerticalAlign = 'top' | 'middle' | 'bottom';

/** A serialized border line containing a style name and optional CSS color. */
export type BorderLine = readonly [style: string, color?: string];

/**
 * Font overrides stored in a cell style.
 *
 * @interface
 */
export type FontStyle = JsonExtensible<{
  /** Font family name. */
  readonly name?: string;
  /** Font size in points. */
  readonly size?: number;
  /** Whether the text is bold. */
  readonly bold?: boolean;
  /** Whether the text is italic. */
  readonly italic?: boolean;
}>;

/**
 * Border lines applied to the four sides of a cell.
 *
 * @interface
 */
export type CellBorders = JsonExtensible<{
  /** Top border line. */
  readonly top?: BorderLine;
  /** Right border line. */
  readonly right?: BorderLine;
  /** Bottom border line. */
  readonly bottom?: BorderLine;
  /** Left border line. */
  readonly left?: BorderLine;
}>;

/**
 * Serializable visual formatting for a cell, row, or column.
 *
 * @interface
 */
export type CellStyle = JsonExtensible<{
  /** Display format identifier, such as a number or date format. */
  readonly format?: string;
  /** Cell background as a CSS color. */
  readonly bgcolor?: string;
  /** Horizontal text alignment. */
  readonly align?: HorizontalAlign;
  /** Vertical text alignment. */
  readonly valign?: VerticalAlign;
  /** Whether text wraps within the cell. */
  readonly textwrap?: boolean;
  /** Whether text has a strike-through decoration. */
  readonly strike?: boolean;
  /** Whether text has an underline decoration. */
  readonly underline?: boolean;
  /** Text color as a CSS color. */
  readonly color?: string;
  /** Font-specific overrides. */
  readonly font?: FontStyle;
  /** Border lines around the cell. */
  readonly border?: CellBorders;
}>;

/**
 * Serializable contents and metadata for one cell.
 *
 * @interface
 */
export type CellData = JsonExtensible<{
  /** Editable source text; formulas retain their leading `=` rather than a rendered value. */
  readonly text?: string;
  /** Zero-based index into the containing sheet's `styles` array. */
  readonly style?: number;
  /** Additional row and column spans for a merge anchor. */
  readonly merge?: readonly [rowSpan: number, columnSpan: number];
  /** Whether users may edit this cell when the sheet is otherwise writable. */
  readonly editable?: boolean;
  /** Whether printed content is visible; `false` preserves style and merged-cell geometry. */
  readonly printable?: boolean;
  /** Optional cached value that may be invalidated when the cell is edited. */
  readonly value?: JsonValue;
}>;

/**
 * Sparse cell data keyed by zero-based decimal column indexes.
 * Omitted columns are empty cells.
 *
 * @interface
 */
export type CellsData = SparseJsonCollection;

/**
 * Serializable properties for one row in a sparse row collection.
 *
 * @interface
 */
export type RowData = JsonExtensible<{
  /** Row height in CSS pixels. */
  readonly height?: number;
  /** Whether the row is hidden. */
  readonly hide?: boolean;
  /** Zero-based index into the containing sheet's `styles` array. */
  readonly style?: number;
  /** Sparse cell data keyed by zero-based decimal column indexes. */
  readonly cells?: CellsData;
}>;

/**
 * Sparse row data keyed by zero-based decimal row indexes.
 * Omitted rows retain their default dimensions and contain no cells.
 *
 * @interface
 */
export type RowsData = SparseJsonCollection<{
  /** Logical row count, including rows omitted from the sparse object. */
  readonly len?: number;
}>;

/**
 * Serializable properties for one column in a sparse column collection.
 *
 * @interface
 */
export type ColumnData = JsonExtensible<{
  /** Column width in CSS pixels. */
  readonly width?: number;
  /** Whether the column is hidden. */
  readonly hide?: boolean;
  /** Zero-based index into the containing sheet's `styles` array. */
  readonly style?: number;
}>;

/**
 * Sparse column data keyed by zero-based decimal column indexes.
 *
 * @interface
 */
export type ColsData = SparseJsonCollection<{
  /** Logical column count, including columns omitted from the sparse object. */
  readonly len?: number;
}>;

/**
 * Serialized validation rule and its target A1 references.
 *
 * @interface
 */
export type ValidationData = JsonExtensible<{
  /** A1 cells or ranges to which the rule applies. */
  readonly refs?: readonly string[];
  /** Validation scope; currently validation is applied per cell. */
  readonly mode?: 'cell';
  /** Kind of value accepted by the rule. */
  readonly type?: ValidationType;
  /** Whether blank cell text fails validation. */
  readonly required?: boolean;
  /** Comparison operator, or `in` for membership in a list. */
  readonly operator?: ValidationOperator | 'in';
  /** JSON-compatible comparison value or list data. */
  readonly value?: JsonValue;
}>;

/**
 * Serialized filter for a single column within an auto-filter range.
 *
 * @interface
 */
export type AutoFilterItemData = JsonExtensible<{
  /** Absolute zero-based column index in the worksheet. */
  readonly ci?: number;
  /** Whether all values or only listed values remain visible. */
  readonly operator?: 'all' | 'in';
  /** Included text values when `operator` is `in`. */
  readonly value?: readonly string[];
}>;

/**
 * Serialized sort applied within an auto-filter range.
 *
 * @interface
 */
export type AutoFilterSortData = JsonExtensible<{
  /** Absolute zero-based column index in the worksheet. */
  readonly ci?: number;
  /** Sort direction. */
  readonly order?: 'asc' | 'desc';
}>;

/**
 * Serialized filter range, column filters, and optional sort state.
 *
 * @interface
 */
export type AutoFilterData = JsonExtensible<{
  /** A1 range containing the filter header and data rows. */
  readonly ref?: string;
  /** Column filters applied within `ref`. */
  readonly filters?: readonly AutoFilterItemData[];
  /** Active sort, or `null` when no sort is applied. */
  readonly sort?: AutoFilterSortData | null;
}>;

/**
 * JSON-compatible representation of one worksheet.
 * Rows, columns, and cells use sparse objects with zero-based decimal keys.
 *
 * @interface
 */
export type SheetData = JsonExtensible<{
  /** Display name shown on the sheet tab. */
  readonly name?: string;
  /** A1 cell whose row and column mark the first unfrozen pane. */
  readonly freeze?: string;
  /** Style table referenced by zero-based style indexes. */
  readonly styles?: readonly CellStyle[];
  /** A1 merge ranges with canonical casing; reversed endpoint order is preserved. */
  readonly merges?: readonly string[];
  /** Sparse row data and logical row count. */
  readonly rows?: RowsData;
  /** Sparse column data and logical column count. */
  readonly cols?: ColsData;
  /** Validation rules and their A1 targets. */
  readonly validations?: readonly ValidationData[];
  /** Filtering and sorting state for one A1 range. */
  readonly autofilter?: AutoFilterData;
}>;

/**
 * Ordered, JSON-compatible collection of worksheets in a workbook.
 * Runtime `SheetId` values are not part of this serialized data.
 */
export type WorkbookData = readonly SheetData[];

/** A single sheet or complete workbook accepted when initializing or controlling the component. */
export type WorkbookInput = SheetData | WorkbookData;
