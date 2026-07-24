import type { JsonValue } from '../../core/types/json';
import type {
  DocumentId,
  DocumentSheetId,
  GroupId,
  ObjectId,
  ResourceId,
  StyleId,
  TemplateId,
  ValidationId,
} from './ids';
import type { SparseCell, SparseCellInput } from './sparse-cells';

export type { JsonValue } from '../../core/types/json';

/** A zero-based cell coordinate. */
export interface CellPoint {
  /** Zero-based row index. */
  readonly row: number;
  /** Zero-based column index. */
  readonly column: number;
}

/** A cell coordinate qualified by its stable sheet identifier. */
export interface DocumentCellAddress extends CellPoint {
  /** Sheet containing the addressed cell. */
  readonly sheetId: DocumentSheetId;
}

/** A normalized inclusive cell range qualified by its sheet identifier. */
export interface DocumentCellRange {
  /** Sheet containing the range. */
  readonly sheetId: DocumentSheetId;
  /** Inclusive top-left coordinate. */
  readonly start: CellPoint;
  /** Inclusive bottom-right coordinate. */
  readonly end: CellPoint;
}

/** A normalized inclusive cell range within one sheet. */
export interface SheetRange {
  /** Inclusive top-left coordinate. */
  readonly start: CellPoint;
  /** Inclusive bottom-right coordinate. */
  readonly end: CellPoint;
}

/** Persistent typed input for a spreadsheet cell. */
export type CellInput =
  | {
      /** Discriminator for an explicitly blank cell. */
      readonly type: 'blank';
    }
  | {
      /** Discriminator for a string cell. */
      readonly type: 'string';
      /** String value, including an explicit empty string. */
      readonly value: string;
    }
  | {
      /** Discriminator for a finite number cell. */
      readonly type: 'number';
      /** Finite numeric value, including Excel date serials. */
      readonly value: number;
    }
  | {
      /** Discriminator for a boolean cell. */
      readonly type: 'boolean';
      /** Boolean value. */
      readonly value: boolean;
    }
  | {
      /** Discriminator for an unevaluated formula cell. */
      readonly type: 'formula';
      /** Formula source text. */
      readonly source: string;
    }
  | {
      /** Discriminator for an application-defined custom cell. */
      readonly type: 'custom';
      /** Namespaced custom cell type. */
      readonly cellType: string;
      /** Version of the custom cell payload schema. */
      readonly schemaVersion: number;
      /** JSON-compatible custom cell payload. */
      readonly value: JsonValue;
    };

/** Persistent data stored for one non-empty cell. */
export interface Cell {
  /** Typed user input for the cell. */
  readonly input: CellInput;
  /** Referenced style registry entry. */
  readonly styleId?: StyleId;
  /** Referenced validation registry entry. */
  readonly validationId?: ValidationId;
  /** Referenced resource entry. */
  readonly resourceId?: ResourceId;
  /** Referenced template entry. */
  readonly templateId?: TemplateId;
  /** Application-defined JSON metadata. */
  readonly metadata?: JsonValue;
  /** Whether users may edit this cell. */
  readonly editable?: boolean;
  /** Whether cell content participates in print output. */
  readonly printable?: boolean;
}

/** Sparse normalized row layout properties. */
export interface SheetRow {
  /** Zero-based row index. */
  readonly index: number;
  /** Row height in CSS pixels. */
  readonly height?: number;
  /** Whether the row is hidden. */
  readonly hidden?: boolean;
  /** Default style inherited by cells in the row. */
  readonly styleId?: StyleId;
}

/** Sparse normalized column layout properties. */
export interface SheetColumn {
  /** Zero-based column index. */
  readonly index: number;
  /** Column width in CSS pixels. */
  readonly width?: number;
  /** Whether the column is hidden. */
  readonly hidden?: boolean;
  /** Default style inherited by cells in the column. */
  readonly styleId?: StyleId;
}

/** Persistent row or column outline definition. */
export interface SheetGroup {
  /** Stable identity used by toggle and ungroup commands. */
  readonly id: GroupId;
  /** Worksheet axis covered by this group. */
  readonly axis: 'row' | 'column';
  /** Inclusive first logical row or column. */
  readonly start: number;
  /** Inclusive last logical row or column. */
  readonly end: number;
  /** One-based nesting depth derived from containment. */
  readonly level: number;
  /** Whether covered entries are hidden in derived presentation. */
  readonly collapsed: boolean;
}

/** One normalized value filter. */
export interface SheetFilterItem {
  /** Absolute zero-based filtered column. */
  readonly column: number;
  /** Whether all values or only listed values remain visible. */
  readonly operator: 'all' | 'in';
  /** Values retained when the operator is `in`. */
  readonly values: readonly string[];
}

/** Normalized sheet filter and sort state. */
export interface SheetFilter {
  /** Inclusive filtered region. */
  readonly range?: SheetRange;
  /** Per-column filter definitions. */
  readonly filters: readonly SheetFilterItem[];
  /** Optional active sort state. */
  readonly sort?: {
    /** Absolute zero-based sorted column. */
    readonly column: number;
    /** Sort direction. */
    readonly direction: 'asc' | 'desc';
  } | null;
}

/** One saved-view comparison over an absolute worksheet column. */
export interface FilterViewPredicate {
  /** Absolute zero-based worksheet column. */
  readonly column: number;
  /** Scalar comparison operator. */
  readonly operator:
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'greaterThanOrEqual'
    | 'lessThan'
    | 'lessThanOrEqual'
    | 'contains';
  /** Fixed comparison value. */
  readonly value: string | number | boolean;
}

/** Persistent saved filter and sort definition; active selection remains session-owned. */
export interface FilterView {
  /** Stable view identifier. */
  readonly id: string;
  /** User-visible view name. */
  readonly name: string;
  /** Header-inclusive source range. */
  readonly range: DocumentCellRange;
  /** Ordered sort definitions. */
  readonly sorts: readonly {
    /** Absolute zero-based worksheet column. */
    readonly column: number;
    /** Sort direction. */
    readonly direction: 'ascending' | 'descending';
  }[];
  /** Predicates combined with logical AND. */
  readonly filters: readonly FilterViewPredicate[];
  /** Persistent document definition or ephemeral session definition. */
  readonly visibility: 'document' | 'session';
}

/** Device-independent object rectangle or cell offset. */
export interface ObjectRect {
  /** Horizontal position. */
  readonly x: number;
  /** Vertical position. */
  readonly y: number;
  /** Non-negative width. */
  readonly width: number;
  /** Non-negative height. */
  readonly height: number;
}

/** Logical positioning modes shared by worksheet objects. */
export type ObjectAnchor =
  | {
      /** Uses absolute worksheet geometry. */
      readonly type: 'absolute';
      /** Fixed object rectangle. */
      readonly rect: ObjectRect;
    }
  | {
      /** Positions a fixed-size object relative to one cell. */
      readonly type: 'one-cell';
      /** Anchor cell. */
      readonly cell: DocumentCellAddress;
      /** Offset from the cell origin. */
      readonly offset: {
        /** Horizontal cell offset. */
        readonly x: number;
        /** Vertical cell offset. */
        readonly y: number;
      };
      /** Fixed object size. */
      readonly size: Pick<ObjectRect, 'width' | 'height'>;
    }
  | {
      /** Sizes an object between two cell markers. */
      readonly type: 'two-cell';
      /** Top-left marker. */
      readonly from: DocumentCellAddress & {
        /** Offset from the top-left cell origin. */
        readonly offset: Pick<ObjectRect, 'x' | 'y'>;
      };
      /** Bottom-right marker. */
      readonly to: DocumentCellAddress & {
        /** Offset from the bottom-right cell origin. */
        readonly offset: Pick<ObjectRect, 'x' | 'y'>;
      };
    };

/** Persistent image, bounded vector shape, or plain-text floating worksheet object. */
export type SheetObject =
  | {
      /** Stable object identifier. */
      readonly id: ObjectId;
      /** Raster image object. */
      readonly kind: 'image';
      /** Logical object position. */
      readonly anchor: ObjectAnchor;
      /** Stable paint order. */
      readonly zIndex: number;
      /** Whether interactive editing is locked. */
      readonly locked: boolean;
      /** Template repetition policy. */
      readonly templateRepeat: 'shared' | 'per-item' | 'forbidden';
      /** Document-owned image resource. */
      readonly resourceId: ResourceId;
      /** Image fitting policy. */
      readonly fit?: 'contain' | 'cover' | 'fill';
      /** Clockwise rotation normalized to the half-open range [0, 360). */
      readonly rotation?: number;
      /** Accessible object label. */
      readonly accessibility: {
        /** Accessible object name. */
        readonly name: string;
        /** Optional accessible object description. */
        readonly description?: string;
      };
    }
  | {
      /** Stable object identifier. */
      readonly id: ObjectId;
      /** Bounded renderer-neutral vector shape. */
      readonly kind: 'shape';
      /** Supported primitive shape. */
      readonly shape: 'rectangle' | 'ellipse' | 'line';
      /** Logical object position. */
      readonly anchor: ObjectAnchor;
      /** Stable paint order. */
      readonly zIndex: number;
      /** Whether interactive editing is locked. */
      readonly locked: boolean;
      /** Template repetition policy. */
      readonly templateRepeat: 'shared' | 'per-item' | 'forbidden';
      /** Clockwise rotation normalized to the half-open range [0, 360). */
      readonly rotation?: number;
      /** Deterministic shape presentation. */
      readonly style: {
        /** Optional fill color; omitted for a transparent interior. */
        readonly fill?: string;
        /** Optional border or line color. */
        readonly stroke?: string;
        /** Optional non-negative border or line width. */
        readonly strokeWidth?: number;
      };
      /** Accessible object label. */
      readonly accessibility: {
        /** Accessible object name. */
        readonly name: string;
        /** Optional accessible object description. */
        readonly description?: string;
      };
    }
  | {
      /** Stable object identifier. */
      readonly id: ObjectId;
      /** Plain-text box object. */
      readonly kind: 'text-box';
      /** Logical object position. */
      readonly anchor: ObjectAnchor;
      /** Stable paint order. */
      readonly zIndex: number;
      /** Whether interactive editing is locked. */
      readonly locked: boolean;
      /** Template repetition policy. */
      readonly templateRepeat: 'shared' | 'per-item' | 'forbidden';
      /** Plain text that is never interpreted as markup. */
      readonly text: string;
      /** Clockwise rotation normalized to the half-open range [0, 360). */
      readonly rotation?: number;
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
      /** Accessible object label. */
      readonly accessibility: {
        /** Accessible object name. */
        readonly name: string;
        /** Optional accessible object description. */
        readonly description?: string;
      };
    };

/** Workbook-tab visibility persisted independently of any output adapter. */
export type WorksheetVisibility = 'visible' | 'hidden' | 'very-hidden';

/** Differential style supported by persistent cell-is conditional rules. */
export interface ConditionalStyle {
  /** Optional text ARGB or RGB color. */
  readonly color?: string;
  /** Optional fill ARGB or RGB color. */
  readonly backgroundColor?: string;
  /** Optional bold emphasis. */
  readonly bold?: boolean;
}

/** Persistent two- or three-color scale over one worksheet range. */
export interface ConditionalColorScale {
  /** Stable color-scale discriminator. */
  readonly type: 'color-scale';
  /** Qualified worksheet range receiving the rule. */
  readonly range: DocumentCellRange;
  /** Minimum-value ARGB or RGB color. */
  readonly minimumColor: string;
  /** Optional 50th-percentile ARGB or RGB color. */
  readonly midpointColor?: string;
  /** Maximum-value ARGB or RGB color. */
  readonly maximumColor: string;
}

/** Persistent formula-backed cell comparison over one worksheet range. */
export interface ConditionalCellRule {
  /** Stable cell-comparison discriminator. */
  readonly type: 'cell-is';
  /** Qualified worksheet range receiving the rule. */
  readonly range: DocumentCellRange;
  /** Allowlisted spreadsheet comparison operator. */
  readonly operator:
    | 'between'
    | 'notBetween'
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual';
  /** First formula in the restricted safe subset, without a leading equals sign. */
  readonly formula: string;
  /** Required second formula for between and notBetween. */
  readonly formula2?: string;
  /** Differential style applied when the comparison matches. */
  readonly style: ConditionalStyle;
}

/** Ordered persistent conditional-format rule. */
export type ConditionalFormat = ConditionalColorScale | ConditionalCellRule;

/** Ordered worksheet with sparse cells and normalized merges. */
export interface Sheet {
  /** Stable opaque sheet identifier. */
  readonly id: DocumentSheetId;
  /** User-visible sheet name. */
  readonly name: string;
  /** Sparse non-empty cells in canonical coordinate order. */
  readonly cells: readonly SparseCell[];
  /** User-defined normalized merge ranges. */
  readonly merges: readonly SheetRange[];
  /** Logical row count when explicitly supplied by the source. */
  readonly rowCount?: number;
  /** Logical column count when explicitly supplied by the source. */
  readonly columnCount?: number;
  /** Sparse row layout in ascending index order. */
  readonly rows: readonly SheetRow[];
  /** Sparse column layout in ascending index order. */
  readonly columns: readonly SheetColumn[];
  /** Canonically ordered persistent outline groups. */
  readonly groups: readonly SheetGroup[];
  /** First unfrozen row and column. */
  readonly freeze?: CellPoint;
  /** Optional normalized filter and sort state. */
  readonly filter?: SheetFilter;
  /** Workbook-tab visibility retained by semantic output adapters. */
  readonly visibility: WorksheetVisibility;
  /** Ordered typed conditional-format rules owned by this worksheet. */
  readonly conditionalFormatting: readonly ConditionalFormat[];
  /** Persistent saved-view definitions; active selection is never stored here. */
  readonly filterViews: readonly FilterView[];
  /** Canonically ordered floating worksheet objects. */
  readonly objects: readonly SheetObject[];
}

/** One stable-ID entry in a JSON-valued registry. */
export interface RegistryEntry<Id extends string> {
  /** Stable opaque entry identifier. */
  readonly id: Id;
  /** JSON-compatible registry value. */
  readonly value: JsonValue;
}

/** Canonically ordered style definitions. */
export type StyleRegistry = readonly RegistryEntry<StyleId>[];
/** Canonically ordered validation definitions. */
export type ValidationRegistry = readonly RegistryEntry<ValidationId>[];

/** Persistent workbook-level interpretation settings. */
export interface WorkbookSettings {
  /** Excel serial-date system used by numeric date cells. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  /** Optional locale hint for consumers. */
  readonly localeHint?: string;
}

/** Ordered sheets, registries, and settings for a document. */
export interface Workbook {
  /** User-ordered worksheets. */
  readonly sheets: readonly Sheet[];
  /** Canonically ordered style registry. */
  readonly styles: StyleRegistry;
  /** Canonically ordered validation registry. */
  readonly validations: ValidationRegistry;
  /** Workbook interpretation settings. */
  readonly settings: WorkbookSettings;
}

/** Print margins expressed in profile-defined units. */
export interface PrintMargins {
  /** Top margin. */
  readonly top: number;
  /** Right margin. */
  readonly right: number;
  /** Bottom margin. */
  readonly bottom: number;
  /** Left margin. */
  readonly left: number;
}

/** Minimal persistent print settings for a template. */
export interface PrintProfile {
  /** Consumer-defined paper size name. */
  readonly paperSize: string;
  /** Page orientation. */
  readonly orientation: 'portrait' | 'landscape';
  /** Page margins. */
  readonly margins: PrintMargins;
  /** Optional finite print scale. */
  readonly scale?: number;
}

/**
 * Reusable TP1 template persisted inside the source document.
 *
 * Binding and print-profile definitions live in the template compiler module,
 * while this document-owned interface is the canonical persistence boundary.
 */
export interface StoredSpreadsheetTemplate {
  /** Stable opaque template identifier. */
  readonly id: TemplateId;
  /** User-visible template name. */
  readonly name: string;
  /** Explicit binding metadata compiled without scanning cell text. */
  readonly bindings: readonly import('../../template/model').TemplateBinding[];
  /** Named deterministic print profiles. */
  readonly printProfiles: readonly import('../../template/model').TemplatePrintProfile[];
}

/** Metadata describing a resource without fetching it. */
export interface ResourceMetadata {
  /** Stable opaque resource identifier. */
  readonly id: ResourceId;
  /** Consumer-defined resource kind. */
  readonly kind: string;
  /** Optional MIME type. */
  readonly mimeType?: string;
  /** Optional unresolved resource URL. */
  readonly url?: string;
  /** Optional declared byte length. */
  readonly byteLength?: number;
  /** Optional JSON-compatible resource metadata. */
  readonly metadata?: JsonValue;
}

/** Canonically ordered persistent resource metadata. */
export interface ResourceStore {
  /** Resource metadata entries. */
  readonly items: readonly ResourceMetadata[];
}

/** Namespaced JSON-compatible extension data. */
export interface ExtensionStore {
  readonly [namespace: string]: JsonValue;
}

/** Immutable persistent Workbook 2.0 document snapshot. */
export interface SpreadsheetDocument {
  /** Workbook schema version, exactly 2. */
  readonly schemaVersion: 2;
  /** Stable opaque document identifier. */
  readonly id: DocumentId;
  /** Workbook content and settings. */
  readonly workbook: Workbook;
  /** User-ordered reusable templates. */
  readonly templates: readonly StoredSpreadsheetTemplate[];
  /** Resource metadata store. */
  readonly resources: ResourceStore;
  /** Namespaced extension data. */
  readonly extensions: ExtensionStore;
}

export interface SheetInput {
  id: string;
  name: string;
  cells: SparseCellInput[];
  merges: { start: CellPoint; end: CellPoint }[];
  rowCount?: number;
  columnCount?: number;
  rows?: { index: number; height?: number; hidden?: boolean; styleId?: string }[];
  columns?: { index: number; width?: number; hidden?: boolean; styleId?: string }[];
  groups?: {
    id: string;
    axis: 'row' | 'column';
    start: number;
    end: number;
    level: number;
    collapsed: boolean;
  }[];
  freeze?: CellPoint;
  filter?: {
    range?: { start: CellPoint; end: CellPoint };
    filters: { column: number; operator: 'all' | 'in'; values: string[] }[];
    sort?: { column: number; direction: 'asc' | 'desc' } | null;
  };
  visibility?: WorksheetVisibility;
  conditionalFormatting?: ConditionalFormat[];
  filterViews?: FilterView[];
  objects?: SheetObject[];
}

export interface SpreadsheetDocumentInput {
  schemaVersion: number;
  id: string;
  workbook: {
    sheets: SheetInput[];
    styles: { id: string; value: JsonValue }[];
    validations: { id: string; value: JsonValue }[];
    settings: {
      dateSystem: 'excel-1900' | 'excel-1904';
      localeHint?: string;
    };
  };
  templates: {
    id: string;
    name: string;
    bindings: import('../../template/model').TemplateBinding[];
    printProfiles: import('../../template/model').TemplatePrintProfile[];
  }[];
  resources: {
    items: {
      id: string;
      kind: string;
      mimeType?: string;
      url?: string;
      byteLength?: number;
      metadata?: JsonValue;
    }[];
  };
  extensions: Record<string, JsonValue>;
}

/** Options for creating an empty Workbook 2.0 document. */
export interface CreateDocumentOptions {
  /** Caller-provided stable document identifier. */
  readonly id?: string;
  /** Caller-provided stable initial sheet identifier. */
  readonly sheetId?: string;
  /** Initial sheet display name. */
  readonly sheetName?: string;
  /** Excel serial-date system. */
  readonly dateSystem?: 'excel-1900' | 'excel-1904';
  /** Optional locale hint. */
  readonly localeHint?: string;
}
