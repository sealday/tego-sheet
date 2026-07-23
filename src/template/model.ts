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
import type { FormulaValue } from '../formula';
import type { FontMetrics } from '../presentation';
import type { PrintDisplayList } from '../print';
import type { CompiledTemplateExpression, TemplateFormatterRegistry } from './expression';
import type {
  DecodedResourceImage,
  ResolvedResourceStore,
  ResourcePurpose,
  ResourceRef,
  ResourceResolverRegistry,
} from './resources';

/** Binding kinds supported by the template compiler. */
export type TemplateBinding =
  | ValueBinding
  | RepeatRowsBinding
  | RepeatColumnsBinding
  | RepeatRangeBinding
  | RepeatPageBinding
  | RepeatSheetBinding
  | ConditionalRangeBinding
  | SubtemplateBinding;

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

/** Explicit floating-object behavior for a repeated structural region. */
export type ObjectRepeatPolicy = 'per-item' | 'shared' | 'forbidden';

/** Stable floating-object reference intersecting a structural region. */
export interface RepeatedObjectRef {
  /** Stable object identity. */
  readonly id: string;
  /** Source anchor transformed with the repeated region. */
  readonly anchor: DocumentCellRange;
  /** Whether the anchor follows structural mappings or remains absolute. */
  readonly anchorMode: 'range' | 'absolute';
  /** Optional content-addressed logical resource used by the object. */
  readonly resourceId?: string;
}

/** Shared structural-repeat source and policy fields. */
export interface AdvancedRepeatBase {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Inclusive source region. */
  readonly range: DocumentCellRange;
  /** Restricted collection expression. */
  readonly source: string;
  /** Empty-collection behavior. */
  readonly empty: 'remove' | 'keep-template-row';
  /** Explicit object-copy behavior. */
  readonly objectPolicy?: ObjectRepeatPolicy;
  /** Compiler-visible object references intersecting this region. */
  readonly objects?: readonly RepeatedObjectRef[];
}

/** Repeats a source range along the column axis. */
export interface RepeatColumnsBinding extends AdvancedRepeatBase {
  /** Horizontal-repeat discriminator. */
  readonly type: 'repeat-columns';
}

/** Repeats a source rectangle along one or both axes. */
export interface RepeatRangeBinding extends AdvancedRepeatBase {
  /** Rectangular-repeat discriminator. */
  readonly type: 'repeat-range';
  /** Axes along which copies are created. */
  readonly axis: 'vertical' | 'horizontal' | 'both';
}

/** Repeats one logical fragment with a hard boundary before every later item. */
export interface RepeatPageBinding extends AdvancedRepeatBase {
  /** Per-item-page discriminator. */
  readonly type: 'repeat-page';
}

/** Clones a source sheet once per item. */
export interface RepeatSheetBinding extends AdvancedRepeatBase {
  /** Per-item-sheet discriminator. */
  readonly type: 'repeat-sheet';
  /** Restricted expression producing the visible sheet name. */
  readonly name: string;
}

/** Inserts one explicitly registered compatible template. */
export interface SubtemplateBinding {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Subtemplate discriminator. */
  readonly type: 'subtemplate';
  /** Target region receiving the child template. */
  readonly range: DocumentCellRange;
  /** Explicitly registered child template identity. */
  readonly templateId: TemplateId;
  /** Restricted expression producing child-template data. */
  readonly source: string;
  /** Explicit object-copy behavior. */
  readonly objectPolicy?: ObjectRepeatPolicy;
  /** Compiler-visible intersecting objects. */
  readonly objects?: readonly RepeatedObjectRef[];
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

/** One validated resource positioned over a template cell range. */
export interface TemplateResourceBinding {
  /** Stable binding identity. */
  readonly id: BindingId;
  /** Cells covered by the resource in template coordinates. */
  readonly target: DocumentCellRange;
  /** Logical resource reference resolved by the render session. */
  readonly resourceId: string;
  /** Deterministic sizing policy inside the target rectangle. */
  readonly fit: 'contain' | 'cover' | 'fill';
}

/** Full template model persisted in `SpreadsheetDocument.templates`. */
export type SpreadsheetTemplate = StoredSpreadsheetTemplate & {
  /** Optional resource placements resolved before layout and painted as overlays. */
  readonly resourceBindings?: readonly TemplateResourceBinding[];
};

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
  | (Omit<RepeatColumnsBinding, 'source'> & {
      /** Parsed collection expression. */
      readonly source: CompiledTemplateExpression;
    })
  | (Omit<RepeatRangeBinding, 'source'> & {
      /** Parsed collection expression. */
      readonly source: CompiledTemplateExpression;
    })
  | (Omit<RepeatPageBinding, 'source'> & {
      /** Parsed collection expression. */
      readonly source: CompiledTemplateExpression;
    })
  | (Omit<RepeatSheetBinding, 'source' | 'name'> & {
      /** Parsed collection expression. */
      readonly source: CompiledTemplateExpression;
      /** Parsed generated-sheet-name expression. */
      readonly name: CompiledTemplateExpression;
    })
  | (Omit<SubtemplateBinding, 'source'> & {
      /** Parsed child-data expression. */
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
  /** Ordered advanced structural containment forest. */
  readonly regionTree?: readonly TemplateRegionNode[];
  /** Explicit compile-time subtemplate registry snapshot. */
  readonly subtemplates?: readonly SpreadsheetTemplate[];
}

/** One immutable node in the validated structural-containment forest. */
export interface TemplateRegionNode {
  /** Structural binding represented by this node. */
  readonly bindingId: BindingId;
  /** Original source region. */
  readonly range: DocumentCellRange;
  /** One-based containment depth. */
  readonly depth: number;
  /** Deterministically ordered fully-contained children. */
  readonly children: readonly TemplateRegionNode[];
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
  /** Maximum expanded columns. */
  readonly maxExpandedColumns?: number;
  /** Maximum generated worksheets. */
  readonly maxGeneratedSheets?: number;
  /** Maximum generated pages. */
  readonly maxPages: number;
  /** Maximum layout wall-clock duration. */
  readonly maxLayoutTimeMs: number;
  /** Maximum structural nesting depth. */
  readonly maxNestingDepth?: number;
  /** Maximum logical resource references. */
  readonly maxResources?: number;
  /** Maximum compressed bytes for one resource. */
  readonly maxResourceBytes?: number;
  /** Maximum compressed bytes for the session. */
  readonly maxTotalResourceBytes?: number;
  /** Maximum concurrent resolver calls. */
  readonly maxResolveConcurrency?: number;
}

/** Additional inputs required to compile TP2 structural templates. */
export interface AdvancedCompileOptions {
  /** Explicit schema-compatible child templates keyed by stable ID. */
  readonly subtemplates: ReadonlyMap<TemplateId, SpreadsheetTemplate>;
  /** Compile and expansion safety budgets. */
  readonly limits: RenderLimits;
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
  /** Explicit logical resources discovered for this render session. */
  readonly resourceRefs?: readonly ResourceRef[];
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
  /** Explicit resolver registry; it never installs a network client implicitly. */
  readonly resourceRegistry?: ResourceResolverRegistry;
  /** Output purpose forwarded to capability-limited resolvers. */
  readonly resourcePurpose?: ResourcePurpose;
  /** Host image decoder used after MIME and quota validation. */
  readonly decodeImage?: (
    bytes: Uint8Array,
    mimeType: string,
    signal: AbortSignal,
  ) => Promise<DecodedResourceImage>;
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
  /** First target-relative column on the page. */
  readonly columnStart: number;
  /** Last target-relative column on the page. */
  readonly columnEnd: number;
}

/** Immutable pagination output shared by preview and output adapters. */
export interface PrintDocument {
  /** Ordered semantic pages. */
  readonly pages: readonly GeneratedPrintPage[];
  /** Exact renderer-neutral display commands for those pages. */
  readonly displayList: PrintDisplayList;
  /** Exact immutable print profile selected for this render. */
  readonly profile: TemplatePrintProfile;
}

/** One calculated value retained for semantic output adapters. */
export interface GeneratedCalculatedCell {
  /** Stable generated-workbook address. */
  readonly address: DocumentCellAddress;
  /** Typed calculated or input value used by presentation. */
  readonly value: FormulaValue;
}

/** Typed two- or three-color conditional formatting for one generated range. */
export interface GeneratedConditionalColorScale {
  /** Stable conditional-format discriminator. */
  readonly type: 'color-scale';
  /** Generated workbook range receiving the rule. */
  readonly range: DocumentCellRange;
  /** Minimum-value ARGB or RGB color. */
  readonly minimumColor: string;
  /** Optional 50th-percentile ARGB or RGB color. */
  readonly midpointColor?: string;
  /** Maximum-value ARGB or RGB color. */
  readonly maximumColor: string;
}

/** Semantic worksheet output settings retained for XLSX translation. */
export interface GeneratedWorksheet {
  /** Stable generated worksheet identity. */
  readonly sheetId: DocumentSheetId;
  /** Workbook-tab visibility. */
  readonly visibility: 'visible' | 'hidden' | 'very-hidden';
  /** Ordered typed conditional-format rules. */
  readonly conditionalFormatting: readonly GeneratedConditionalColorScale[];
}

/** Atomic render artifact consumed by every TP1 output surface. */
export interface GeneratedDocument {
  /** Expanded and recalculated semantic workbook. */
  readonly workbook: SpreadsheetDocument['workbook'];
  /** Canonically ordered calculated values for generated sparse cells. */
  readonly calculatedCells: readonly GeneratedCalculatedCell[];
  /** Ordered semantic worksheet output settings. */
  readonly worksheets: readonly GeneratedWorksheet[];
  /** Shared immutable print pages and commands. */
  readonly print: PrintDocument;
  /** Session-owned, content-addressed resolved resource store. */
  readonly resources: ResolvedResourceStore;
  /** Generated floating-object anchors after structural expansion. */
  readonly objects: readonly import('./expand').StructuralObjectMapping[];
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
