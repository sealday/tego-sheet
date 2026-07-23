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

export interface CellPoint {
  readonly row: number;
  readonly column: number;
}

export interface DocumentCellAddress extends CellPoint {
  readonly sheetId: DocumentSheetId;
}

export interface DocumentCellRange {
  readonly sheetId: DocumentSheetId;
  readonly start: CellPoint;
  readonly end: CellPoint;
}

export interface SheetRange {
  readonly start: CellPoint;
  readonly end: CellPoint;
}

export type CellInput =
  | { readonly type: 'blank' }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'boolean'; readonly value: boolean }
  | { readonly type: 'formula'; readonly source: string }
  | {
      readonly type: 'custom';
      readonly cellType: string;
      readonly schemaVersion: number;
      readonly value: JsonValue;
    };

export interface Cell {
  readonly input: CellInput;
  readonly styleId?: StyleId;
  readonly validationId?: ValidationId;
  readonly resourceId?: ResourceId;
  readonly templateId?: TemplateId;
  readonly metadata?: JsonValue;
}

export interface Sheet {
  readonly id: DocumentSheetId;
  readonly name: string;
  readonly cells: readonly SparseCell[];
  readonly merges: readonly SheetRange[];
}

export interface RegistryEntry<Id extends string> {
  readonly id: Id;
  readonly value: JsonValue;
}

export type StyleRegistry = readonly RegistryEntry<StyleId>[];
export type ValidationRegistry = readonly RegistryEntry<ValidationId>[];

export interface WorkbookSettings {
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  readonly localeHint?: string;
}

export interface Workbook {
  readonly sheets: readonly Sheet[];
  readonly styles: StyleRegistry;
  readonly validations: ValidationRegistry;
  readonly settings: WorkbookSettings;
}

export interface PrintMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

export interface PrintProfile {
  readonly paperSize: string;
  readonly orientation: 'portrait' | 'landscape';
  readonly margins: PrintMargins;
  readonly scale?: number;
}

export interface SpreadsheetTemplate {
  readonly id: TemplateId;
  readonly name: string;
  readonly sheetId: DocumentSheetId;
  readonly range?: DocumentCellRange;
  readonly printProfile: PrintProfile;
}

export interface ResourceMetadata {
  readonly id: ResourceId;
  readonly kind: string;
  readonly mimeType?: string;
  readonly url?: string;
  readonly byteLength?: number;
  readonly metadata?: JsonValue;
}

export interface ResourceStore {
  readonly items: readonly ResourceMetadata[];
}

export interface ExtensionStore {
  readonly [namespace: string]: JsonValue;
}

export interface SpreadsheetDocument {
  readonly schemaVersion: 2;
  readonly id: DocumentId;
  readonly workbook: Workbook;
  readonly templates: readonly SpreadsheetTemplate[];
  readonly resources: ResourceStore;
  readonly extensions: ExtensionStore;
}

export interface SheetInput {
  id: string;
  name: string;
  cells: SparseCellInput[];
  merges: { start: CellPoint; end: CellPoint }[];
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

export interface CreateDocumentOptions {
  readonly id?: string;
  readonly sheetId?: string;
  readonly sheetName?: string;
  readonly dateSystem?: 'excel-1900' | 'excel-1904';
  readonly localeHint?: string;
}
