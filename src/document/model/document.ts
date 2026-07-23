import type { JsonValue } from '../../core/types/json';
import type {
  DocumentId,
  DocumentSheetId,
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
  /** First unfrozen row and column. */
  readonly freeze?: CellPoint;
  /** Optional normalized filter and sort state. */
  readonly filter?: SheetFilter;
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

/** Reusable document template targeting a sheet and optional range. */
export interface StoredSpreadsheetTemplate {
  /** Stable opaque template identifier. */
  readonly id: TemplateId;
  /** User-visible template name. */
  readonly name: string;
  /** Stable identifier of the target sheet. */
  readonly sheetId: DocumentSheetId;
  /** Optional target range on the same sheet. */
  readonly range?: DocumentCellRange;
  /** Persistent print profile. */
  readonly printProfile: PrintProfile;
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
  freeze?: CellPoint;
  filter?: {
    range?: { start: CellPoint; end: CellPoint };
    filters: { column: number; operator: 'all' | 'in'; values: string[] }[];
    sort?: { column: number; direction: 'asc' | 'desc' } | null;
  };
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
    sheetId: string;
    range?: {
      sheetId: string;
      start: CellPoint;
      end: CellPoint;
    };
    printProfile: {
      paperSize: string;
      orientation: 'portrait' | 'landscape';
      margins: { top: number; right: number; bottom: number; left: number };
      scale?: number;
    };
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
