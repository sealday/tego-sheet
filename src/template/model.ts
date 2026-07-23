import type {
  BindingId,
  Diagnostic,
  DocumentCellAddress,
  DocumentCellRange,
  DocumentSheetId,
  SpreadsheetDocument,
  StoredSpreadsheetTemplate,
  TemplateId,
} from '../document';
import type { FontMetrics } from '../presentation';
import type { PrintDisplayList } from '../print';
import type { CompiledTemplateExpression, TemplateFormatterRegistry } from './expression';

/** Binding kinds supported by the TP1 compiler. */
export type TemplateBinding = ValueBinding | RepeatRowsBinding | ConditionalRangeBinding;

/** Writes one resolved scalar into one cell. */
export interface ValueBinding {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Scalar binding discriminator. */
  readonly type: 'value';
  /** Cell receiving the resolved value. */
  readonly target: DocumentCellAddress;
  /** Restricted expression source. */
  readonly expression: string;
  /** Optional registered formatter identifier. */
  readonly formatter?: string;
}

/** Repeats a row range once for every resolved collection item. */
export interface RepeatRowsBinding {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Vertical-repeat discriminator. */
  readonly type: 'repeat-rows';
  /** Inclusive source rows copied for every item. */
  readonly range: DocumentCellRange;
  /** Restricted collection expression. */
  readonly source: string;
  /** Empty-collection behavior. */
  readonly empty: 'remove' | 'keep-template-row';
  /** Per-item pagination preference. */
  readonly pageBreak: 'auto' | 'before-each-item';
}

/** Removes a range when its predicate resolves false. */
export interface ConditionalRangeBinding {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Conditional binding discriminator. */
  readonly type: 'conditional-range';
  /** Inclusive conditional range. */
  readonly range: DocumentCellRange;
  /** Restricted boolean expression. */
  readonly when: string;
}

/** One explicit sheet or range selection submitted to pagination. */
export type PrintTarget =
  | {
      /** Whole-sheet target discriminator. */
      readonly type: 'sheet';
      /** Stable target sheet. */
      readonly sheetId: DocumentSheetId;
    }
  | {
      /** Single-range target discriminator. */
      readonly type: 'range';
      /** Target range. */
      readonly range: DocumentCellRange;
    }
  | {
      /** Ordered multi-range target discriminator. */
      readonly type: 'ranges';
      /** Ranges emitted in declaration order with page boundaries between them. */
      readonly ranges: readonly DocumentCellRange[];
    };

/** Standard or explicit device-independent paper geometry. */
export type PaperDefinition =
  | {
      /** Standard paper name. */
      readonly type: 'A4' | 'A5' | 'Letter';
    }
  | {
      /** Custom paper discriminator. */
      readonly type: 'custom';
      /** Custom paper width. */
      readonly width: number;
      /** Custom paper height. */
      readonly height: number;
    };

/** Explicit print scaling strategy. */
export type PrintScale =
  | {
      /** Fixed-scale discriminator. */
      readonly type: 'fixed';
      /** Positive fixed scale. */
      readonly value: number;
    }
  | {
      /** Width-fitting discriminator. */
      readonly type: 'fit-width';
      /** Requested horizontal page count. */
      readonly pages: number;
    }
  | {
      /** Whole-target page fitting discriminator. */
      readonly type: 'fit-page';
    };

/** Deterministic page geometry used by pagination. */
export interface PageSetup {
  /** Paper geometry. */
  readonly paper: PaperDefinition;
  /** Paper orientation. */
  readonly orientation: 'portrait' | 'landscape';
  /** Non-negative page margins. */
  readonly margins: {
    /** Top margin. */
    readonly top: number;
    /** Right margin. */
    readonly right: number;
    /** Bottom margin. */
    readonly bottom: number;
    /** Left margin. */
    readonly left: number;
  };
  /** Scale policy. */
  readonly scale: PrintScale;
}

/** Manual horizontal page break. */
export interface PageBreak {
  /** Sheet containing the break. */
  readonly sheetId: DocumentSheetId;
  /** Source row which begins the next page. */
  readonly beforeRow: number;
}

/** Static or template-data page band slots. */
export interface PageBand {
  /** Left-aligned band source. */
  readonly left?: string;
  /** Centered band source. */
  readonly center?: string;
  /** Right-aligned band source. */
  readonly right?: string;
}

/** Named collection of print targets and pagination settings. */
export interface TemplatePrintProfile {
  /** Stable profile identity. */
  readonly id: string;
  /** User-visible profile name. */
  readonly name: string;
  /** Ordered output targets. */
  readonly targets: readonly PrintTarget[];
  /** Page geometry. */
  readonly page: PageSetup;
  /** Optional repeated title rows. */
  readonly repeatRows?: DocumentCellRange;
  /** Optional repeated title columns. */
  readonly repeatColumns?: DocumentCellRange;
  /** Explicit page breaks. */
  readonly manualBreaks: readonly PageBreak[];
  /** Optional header slots. */
  readonly header?: PageBand;
  /** Optional footer slots. */
  readonly footer?: PageBand;
  /** Whether cell gridlines are drawn. */
  readonly showGridlines: boolean;
  /** Whether row and column headings are drawn. */
  readonly showHeadings: boolean;
}

/** Full TP1 template model persisted in `SpreadsheetDocument.templates`. */
export type SpreadsheetTemplate = StoredSpreadsheetTemplate;

/** Binding IR with every expression parsed into safe nodes. */
export type TemplateIRBinding =
  | (Omit<ValueBinding, 'expression'> & {
      /** Parsed scalar expression. */
      readonly expression: CompiledTemplateExpression;
    })
  | (Omit<RepeatRowsBinding, 'source'> & {
      /** Parsed collection expression. */
      readonly source: CompiledTemplateExpression;
    })
  | (Omit<ConditionalRangeBinding, 'when'> & {
      /** Parsed predicate expression. */
      readonly when: CompiledTemplateExpression;
    });

/** Immutable output of template parsing and structural validation. */
export interface TemplateIR {
  /** Normalized source template. */
  readonly template: SpreadsheetTemplate;
  /** Parsed bindings. */
  readonly bindings: readonly TemplateIRBinding[];
  /** Validated print profiles. */
  readonly profiles: readonly TemplatePrintProfile[];
}

/** Reusable compiler output bound to one canonical source document hash. */
export interface CompiledTemplate {
  /** Compiled template identity. */
  readonly templateId: TemplateId;
  /** Canonical source document hash. */
  readonly sourceDocumentHash: string;
  /** Immutable source snapshot used to create isolated render sessions. */
  readonly sourceDocument: SpreadsheetDocument;
  /** Safe immutable template IR. */
  readonly ir: TemplateIR;
  /** Non-fatal compilation diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Compiler contract version. */
  readonly compilerVersion: string;
}

/** Atomic result of compiling a template. */
export interface CompilationResult {
  /** Compiled output, absent when any error was found. */
  readonly template?: CompiledTemplate;
  /** Aggregated compilation diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Whether diagnostics contain an error. */
  readonly hasErrors: boolean;
}

/** Synchronous resource budgets enforced by TP1 rendering. */
export interface RenderLimits {
  /** Maximum sparse expanded cells. */
  readonly maxExpandedCells: number;
  /** Maximum expanded rows. */
  readonly maxExpandedRows: number;
  /** Maximum generated pages. */
  readonly maxPages: number;
  /** Maximum layout wall-clock duration. */
  readonly maxLayoutTimeMs: number;
}

/** Explicit immutable input to one template render session. */
export interface RenderRequest {
  /** Precompiled template. */
  readonly template: CompiledTemplate;
  /** Hash of the caller's current document snapshot. */
  readonly currentDocumentHash: string;
  /** Read-only host template data. */
  readonly data: unknown;
  /** Selected print profile identity. */
  readonly profileId: string;
  /** Missing-value policy. */
  readonly missingValue: 'error' | 'warning-and-blank';
  /** Optional cancellation signal propagated through the pipeline. */
  readonly signal?: AbortSignal;
  /** Optional host limit overrides. */
  readonly limits?: Partial<RenderLimits>;
}

/** Deterministic host capabilities required by rendering. */
export interface RenderEnvironment {
  /** BCP 47 formatting locale. */
  readonly locale: string;
  /** IANA formatting time zone. */
  readonly timeZone: string;
  /** Excel serial-date system. */
  readonly dateSystem: 'excel-1900' | 'excel-1904';
  /** Explicit render clock. */
  readonly clock: Date;
  /** Deterministic font measurements. */
  readonly fontMetrics: FontMetrics;
  /** Explicit pure formatter registry. */
  readonly formatters?: TemplateFormatterRegistry;
}

/** Semantic identity and geometry for one generated page. */
export interface GeneratedPrintPage {
  /** Stable page identity. */
  readonly id: string;
  /** Zero-based page index. */
  readonly index: number;
  /** Source target identity. */
  readonly targetId: string;
  /** Page width. */
  readonly width: number;
  /** Page height. */
  readonly height: number;
  /** First target-relative row on the page. */
  readonly rowStart: number;
  /** Last target-relative row on the page. */
  readonly rowEnd: number;
}

/** Immutable pagination output shared by preview and output adapters. */
export interface PrintDocument {
  /** Ordered semantic pages. */
  readonly pages: readonly GeneratedPrintPage[];
  /** Exact renderer-neutral display commands for those pages. */
  readonly displayList: PrintDisplayList;
}

/** Atomic render artifact consumed by every TP1 output surface. */
export interface GeneratedDocument {
  /** Expanded and recalculated semantic workbook. */
  readonly workbook: SpreadsheetDocument['workbook'];
  /** Shared immutable print pages and commands. */
  readonly print: PrintDocument;
  /** TP1 empty resource store reserved for the resource pipeline. */
  readonly resources: Readonly<Record<string, never>>;
  /** Ordered render diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Deterministic render metadata. */
  readonly metadata: {
    /** Template used for rendering. */
    readonly templateId: TemplateId;
    /** Print profile used for rendering. */
    readonly profileId: string;
    /** Source hash validated before rendering. */
    readonly sourceDocumentHash: string;
    /** Formatting locale. */
    readonly locale: string;
    /** Formatting time zone. */
    readonly timeZone: string;
    /** Explicit clock serialized as ISO text. */
    readonly generatedAt: string;
  };
}

/** Atomic render result with no partial document on error. */
export interface RenderResult {
  /** Complete generated artifact, absent on error. */
  readonly document?: GeneratedDocument;
  /** Ordered render diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
}
