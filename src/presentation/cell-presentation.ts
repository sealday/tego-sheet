import type { Diagnostic, DocumentCellAddress } from '../document';
import type { FormulaValue } from '../formula';

/** Fully resolved, renderer-independent cell style. */
export interface ResolvedStyle {
  /** Foreground text color. */
  readonly color: string;
  /** Cell fill color. */
  readonly backgroundColor: string;
  /** Resolved font family name. */
  readonly fontFamily: string;
  /** Font size in presentation units. */
  readonly fontSize: number;
  /** Whether the font uses bold weight. */
  readonly bold: boolean;
  /** Whether the font uses italic style. */
  readonly italic: boolean;
  /** Horizontal text alignment. */
  readonly horizontalAlign: 'left' | 'center' | 'right';
  /** Vertical text alignment. */
  readonly verticalAlign: 'top' | 'middle' | 'bottom';
  /** Whether text may wrap across lines. */
  readonly wrap: boolean;
  /** Optional deterministic number-format code. */
  readonly numberFormat?: string;
  /** Whether text is underlined. */
  readonly underline?: boolean;
  /** Whether text is struck through. */
  readonly strike?: boolean;
  /** Optional resolved border definitions. */
  readonly border?: Readonly<
    Partial<Record<'top' | 'right' | 'bottom' | 'left', readonly [style: string, color?: string]>>
  >;
}

/** Resolved validation state shared by visual and semantic renderers. */
export type PresentationValidation = PresentationValid | PresentationProblem;

/** Presentation state for a cell without a validation problem. */
export interface PresentationValid {
  /** Stable valid-state discriminator. */
  readonly status: 'valid';
}

/** Presentation state for a warning or error. */
export interface PresentationProblem {
  /** Stable problem severity. */
  readonly status: 'warning' | 'error';
  /** Accessible user-facing problem description. */
  readonly message: string;
}

/** Non-mutating annotation attached to a presented cell. */
export interface PresentationAnnotation {
  /** Stable annotation kind. */
  readonly kind: string;
  /** Plain-text annotation description. */
  readonly text: string;
}

/** Immutable semantic output consumed by Canvas, DOM and print adapters. */
export interface CellPresentation {
  /** Stable Workbook 2.0 cell address. */
  readonly address: DocumentCellAddress;
  /** Typed calculated or input value. */
  readonly value: FormulaValue;
  /** Deterministically formatted display text. */
  readonly formattedText: string;
  /** Fully resolved immutable style. */
  readonly style: ResolvedStyle;
  /** Resolved validation state. */
  readonly validation: PresentationValidation;
  /** Ordered non-mutating annotations. */
  readonly annotations: readonly PresentationAnnotation[];
  /** Channel-independent visibility and print participation. */
  readonly visibility: {
    /** Whether row or column layout hides the cell. */
    readonly hidden: boolean;
    /** Whether the cell participates in print output. */
    readonly printable: boolean;
  };
  /** Semantic text and state exposed by accessibility adapters. */
  readonly accessibility: {
    /** Accessible cell name. */
    readonly label: string;
    /** Optional accessible error or annotation description. */
    readonly description?: string;
    /** Whether the cell rejects user edits. */
    readonly readOnly: boolean;
    /** Whether the value or validation state is erroneous. */
    readonly invalid: boolean;
    /** Optional semantic role supplied by a built-in cell type. */
    readonly role?: 'text' | 'checkbox' | 'combobox';
    /** Checked state for checkbox presentations. */
    readonly checked?: boolean;
  };
  /** Recoverable presentation diagnostics. */
  readonly diagnostics?: readonly Diagnostic[];
}

/** Stable document-derived cell lookup used by every renderer. */
export interface PresentationResolver {
  /** Resolves one immutable cell presentation for the configured revision tuple. */
  resolve(address: DocumentCellAddress): CellPresentation;
}
