import type { CellAddress, CellRange, Selection, SheetId } from './coordinates';

/** Category of a committed workbook mutation. */
export type WorkbookChangeKind =
  | 'cell'
  | 'style'
  | 'structure'
  | 'merge'
  | 'clipboard'
  | 'autofill'
  | 'filter'
  | 'validation'
  | 'sheet'
  | 'history';

/** User interaction or imperative API that initiated a change. */
export type ChangeSource =
  | 'keyboard'
  | 'pointer'
  | 'touch'
  | 'toolbar'
  | 'sheet-tabs'
  | 'context-menu'
  | 'clipboard'
  | 'ref';

/** Metadata emitted with the complete workbook after a committed mutation. */
export interface WorkbookChange {
  /** Stable identifier shared by callbacks describing the same commit. */
  readonly id: string;
  /** Category of mutation that was committed. */
  readonly kind: WorkbookChangeKind;
  /** Interaction surface that initiated the mutation. */
  readonly source: ChangeSource;
  /** Worksheet changed by the mutation. */
  readonly sheet: SheetId;
  /** Inclusive affected range when the mutation has a cell range. */
  readonly range?: CellRange;
}

/** Details emitted after a cell text edit commits. */
export interface CellEditEvent {
  /** Identifier matching the corresponding workbook change. */
  readonly changeId: string;
  /** Zero-based address of the edited cell. */
  readonly address: CellAddress;
  /** Cell text immediately before the commit. */
  readonly previousText: string;
  /** Cell text after the commit. */
  readonly text: string;
  /** Interaction surface that initiated the edit. */
  readonly source: ChangeSource;
}

/** Details emitted after an internal or external paste commits. */
export interface PasteEvent {
  /** Identifier matching the corresponding workbook change. */
  readonly changeId: string;
  /** Whether values came from this component or the system clipboard. */
  readonly source: 'internal' | 'external';
  /** Original selection for an internal copy or cut. */
  readonly sourceSelection?: Selection;
  /** Final destination selection after clipping or expansion. */
  readonly target: Selection;
  /** Rectangular matrix of text values written by the paste. */
  readonly values: readonly (readonly string[])[];
}

/**
 * Details emitted when worksheet activation is requested.
 * Imperative ref activation emits even when the requested worksheet is already active.
 */
export interface ActiveSheetChangeEvent {
  /** Newly active worksheet. */
  readonly sheet: SheetId;
  /** Zero-based position of the worksheet in the workbook. */
  readonly index: number;
  /** Interaction surface that activated the worksheet. */
  readonly source: 'sheet-tabs' | 'keyboard' | 'ref';
}

/** Nested per-instance message dictionary used to localize component chrome. */
export interface LocaleMessages {
  /** A translated message or nested message group keyed by locale path segment. */
  readonly [key: string]: string | LocaleMessages;
}

/** Locale identifier and message dictionary supplied to one component instance. */
export interface LocaleDefinition {
  /** Locale identifier, such as `en` or `zh-CN`. */
  readonly id: string;
  /** Messages used by the component's built-in chrome. */
  readonly messages: LocaleMessages;
}
