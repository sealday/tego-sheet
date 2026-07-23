import type {
  BindingId,
  Diagnostic,
  DocumentCellAddress,
  DocumentCellRange,
  DocumentSheetId,
  SpreadsheetDocument,
  TemplateId,
} from '../document';
import type { FontMetrics } from '../presentation';
import type { PrintDisplayList } from '../print';
import type { CompiledTemplateExpression, TemplateFormatterRegistry } from './expression';

export type TemplateBinding = ValueBinding | RepeatRowsBinding | ConditionalRangeBinding;

export interface ValueBinding {
  readonly id: BindingId;
  readonly type: 'value';
  readonly target: DocumentCellAddress;
  readonly expression: string;
  readonly formatter?: string;
}

export interface RepeatRowsBinding {
  readonly id: BindingId;
  readonly type: 'repeat-rows';
  readonly range: DocumentCellRange;
  readonly source: string;
  readonly empty: 'remove' | 'keep-template-row';
  readonly pageBreak: 'auto' | 'before-each-item';
}

export interface ConditionalRangeBinding {
  readonly id: BindingId;
  readonly type: 'conditional-range';
  readonly range: DocumentCellRange;
  readonly when: string;
}

export type PrintTarget =
  | { readonly type: 'sheet'; readonly sheetId: DocumentSheetId }
  | { readonly type: 'range'; readonly range: DocumentCellRange }
  | { readonly type: 'ranges'; readonly ranges: readonly DocumentCellRange[] };

export type PaperDefinition =
  | { readonly type: 'A4' | 'A5' | 'Letter' }
  | { readonly type: 'custom'; readonly width: number; readonly height: number };

export type PrintScale =
  | { readonly type: 'fixed'; readonly value: number }
  | { readonly type: 'fit-width'; readonly pages: number }
  | { readonly type: 'fit-page' };

export interface PageSetup {
  readonly paper: PaperDefinition;
  readonly orientation: 'portrait' | 'landscape';
  readonly margins: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly scale: PrintScale;
}

export interface PageBreak {
  readonly sheetId: DocumentSheetId;
  readonly beforeRow: number;
}

export interface PageBand {
  readonly left?: string;
  readonly center?: string;
  readonly right?: string;
}

export interface TemplatePrintProfile {
  readonly id: string;
  readonly name: string;
  readonly targets: readonly PrintTarget[];
  readonly page: PageSetup;
  readonly repeatRows?: DocumentCellRange;
  readonly repeatColumns?: DocumentCellRange;
  readonly manualBreaks: readonly PageBreak[];
  readonly header?: PageBand;
  readonly footer?: PageBand;
  readonly showGridlines: boolean;
  readonly showHeadings: boolean;
}

/** Full TP1 template model shared by the headless SDK and React editor. */
export interface SpreadsheetTemplate {
  readonly id: TemplateId;
  readonly name: string;
  readonly bindings: readonly TemplateBinding[];
  readonly printProfiles: readonly TemplatePrintProfile[];
}

export type TemplateIRBinding =
  | (Omit<ValueBinding, 'expression'> & { readonly expression: CompiledTemplateExpression })
  | (Omit<RepeatRowsBinding, 'source'> & { readonly source: CompiledTemplateExpression })
  | (Omit<ConditionalRangeBinding, 'when'> & { readonly when: CompiledTemplateExpression });

export interface TemplateIR {
  readonly template: SpreadsheetTemplate;
  readonly bindings: readonly TemplateIRBinding[];
  readonly profiles: readonly TemplatePrintProfile[];
}

export interface CompiledTemplate {
  readonly templateId: TemplateId;
  readonly sourceDocumentHash: string;
  readonly sourceDocument: SpreadsheetDocument;
  readonly ir: TemplateIR;
  readonly diagnostics: readonly Diagnostic[];
  readonly compilerVersion: string;
}

export interface CompilationResult {
  readonly template?: CompiledTemplate;
  readonly diagnostics: readonly Diagnostic[];
  readonly hasErrors: boolean;
}

export interface RenderLimits {
  readonly maxExpandedCells: number;
  readonly maxExpandedRows: number;
  readonly maxPages: number;
  readonly maxLayoutTimeMs: number;
}

export interface RenderRequest {
  readonly template: CompiledTemplate;
  readonly currentDocumentHash: string;
  readonly data: unknown;
  readonly profileId: string;
  readonly missingValue: 'error' | 'warning-and-blank';
  readonly signal?: AbortSignal;
  readonly limits?: Partial<RenderLimits>;
}

export interface RenderEnvironment {
  readonly locale: string;
  readonly timeZone: string;
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  readonly clock: Date;
  readonly fontMetrics: FontMetrics;
  readonly formatters?: TemplateFormatterRegistry;
}

export interface GeneratedPrintPage {
  readonly id: string;
  readonly index: number;
  readonly targetId: string;
  readonly width: number;
  readonly height: number;
  readonly rowStart: number;
  readonly rowEnd: number;
}

export interface PrintDocument {
  readonly pages: readonly GeneratedPrintPage[];
  readonly displayList: PrintDisplayList;
}

export interface GeneratedDocument {
  readonly workbook: SpreadsheetDocument['workbook'];
  readonly print: PrintDocument;
  readonly resources: Readonly<Record<string, never>>;
  readonly diagnostics: readonly Diagnostic[];
  readonly metadata: {
    readonly templateId: TemplateId;
    readonly profileId: string;
    readonly sourceDocumentHash: string;
    readonly locale: string;
    readonly timeZone: string;
    readonly generatedAt: string;
  };
}

export interface RenderResult {
  readonly document?: GeneratedDocument;
  readonly diagnostics: readonly Diagnostic[];
}
