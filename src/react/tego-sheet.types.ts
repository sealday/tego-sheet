import type { CSSProperties } from 'react';
import type {
  ActiveSheetChangeEvent,
  CellAddress,
  CellData,
  CellEditEvent,
  CellStyle,
  LocaleDefinition,
  PasteEvent,
  Selection,
  SheetId,
  SheetOptions,
  TegoSheetError,
  ValidationResult,
  WorkbookChange,
} from '../core';
import type { SpreadsheetDocument } from '../document';
import type { RenderEnvironment, SpreadsheetTemplate } from '../template';
import type {
  ValidationEngineOptions,
  ValidationResult as AdvancedValidationResult,
} from '../validation';
import type { SheetTabsRenderer, ToolbarRenderer } from '../ui/slot-types';
import type { PermissionStore } from '../integrations/permission';
import type { PersistenceSession } from '../integrations/persistence';
import type { CommentPrintPolicy, CommentStore } from '../integrations/comments';
import type { CollaborationSession, PresenceStore } from '../integrations/collaboration';

export type { SheetTabsRenderer, ToolbarRenderer } from '../ui/slot-types';

/** Commit, selection, and error notifications emitted by a mounted spreadsheet. */
export interface TegoSheetCallbacks {
  /**
   * Runs after a mutation commits, before any edit, paste, or resulting selection callback.
   * `value` is the complete workbook snapshot and `change` is its {@link WorkbookChange} metadata.
   * External controlled-value replacements do not emit this callback.
   */
  readonly onDocumentChange?: (nextDocument: SpreadsheetDocument, change: WorkbookChange) => void;
  /** Runs after a worksheet is activated and reports its identifier, index, and source. */
  readonly onActiveSheetChange?: (event: ActiveSheetChangeEvent) => void;
  /** Runs after the active selection changes, including selection changes caused by a commit. */
  readonly onSelectionChange?: (selection: Selection) => void;
  /** Runs after `onDocumentChange` when a cell text edit commits. */
  readonly onCellEdit?: (event: CellEditEvent) => void;
  /** Runs after `onDocumentChange` when an internal or external paste commits. */
  readonly onPaste?: (event: PasteEvent) => void;
  /** Runs when the component handles an operation failure and exposes its structured payload. */
  readonly onError?: (error: TegoSheetError) => void;
  /** Receives compiler and render diagnostics from template surfaces. */
  readonly onDiagnostics?: (diagnostics: readonly import('../document').Diagnostic[]) => void;
  /** Receives immutable template edits from the template property panel. */
  readonly onTemplateChange?: (template: SpreadsheetTemplate) => void;
}

/**
 * Props for the `TegoSheet` React component.
 *
 * @remarks
 * Choose controlled `document` or uncontrolled `defaultDocument` when mounting. Supplying both, or
 * switching a mounted instance between those ownership modes, throws a `TegoSheetException`.
 *
 * @inline
 */
interface TegoSheetSharedProps extends TegoSheetCallbacks {
  /** Active product surface; defaults to the ordinary spreadsheet editor. */
  readonly mode?: 'spreadsheet' | 'template' | 'preview';
  /** Full TP1 template model used by template and preview modes. */
  readonly template?: SpreadsheetTemplate;
  /** Print profile selected by template and preview modes; defaults to the first profile. */
  readonly activePrintProfileId?: string;
  /** Receives print profile selection changes from the template designer. */
  readonly onActivePrintProfileChange?: (profileId: string) => void;
  /** Sample structured data resolved by preview mode. */
  readonly sampleData?: unknown;
  /** Deterministic preview environment; required when mode is `preview`. */
  readonly renderEnvironment?: RenderEnvironment;
  /** Zero-based worksheet index selected on mount. */
  readonly initialActiveSheetIndex?: number;
  /** Disables workbook mutations while preserving navigation, selection, copy, and printing. */
  readonly readOnly?: boolean;
  /** Optional live host permission store; installed stores default every protected action to deny. */
  readonly permissionStore?: PermissionStore;
  /** Optional host-owned persistence session attached to committed workbook transactions. */
  readonly persistenceSession?: PersistenceSession;
  /** Optional host comment store used to project semantic cell/range markers. */
  readonly commentStore?: CommentStore;
  /** Explicit comment print projection policy; defaults to excluding comments. */
  readonly commentPrintPolicy?: CommentPrintPolicy;
  /** Optional ephemeral collaboration presence projected as semantic participants. */
  readonly presenceStore?: PresenceStore;
  /** Optional collaboration connection session projected as accessible live status. */
  readonly collaborationSession?: CollaborationSession;
  /** Per-instance locale identifier and message dictionary for built-in chrome. */
  readonly locale?: LocaleDefinition;
  /** Per-instance worksheet behavior and layout settings. */
  readonly options?: SheetOptions;
  /** Restricted resolver/formula capabilities used by cell-owned validation rules. */
  readonly validationEngine?: ValidationEngineOptions;
  /** Explicit confirmation gate for warning-mode Workbook 2.0 validation rules. */
  readonly confirmValidationWarning?: (
    result: AdvancedValidationResult,
  ) => boolean | Promise<boolean>;
  /** Uses the default toolbar, hides it, or replaces it with a custom renderer. */
  readonly toolbar?: 'default' | false | ToolbarRenderer;
  /** Uses the default sheet tabs, hides them, or replaces them with a custom renderer. */
  readonly sheetTabs?: 'default' | false | SheetTabsRenderer;
  /** Additional class name appended to the root spreadsheet element. */
  readonly className?: string;
  /** Inline styles applied to the root spreadsheet element. */
  readonly style?: CSSProperties;
}

/** @inline */
type TegoSheetOwnership =
  | {
      /** Controlled schema 2 document owned by the parent. */
      readonly ['document']: SpreadsheetDocument;
      /** Excluded when the component is controlled. */
      readonly defaultDocument?: never;
    }
  | {
      /** Excluded when the component is uncontrolled. */
      readonly ['document']?: never;
      /** Schema 2 document read once when mounting an uncontrolled spreadsheet. */
      readonly defaultDocument: SpreadsheetDocument;
    };

/**
 * Props for the `TegoSheet` React component.
 *
 * @remarks
 * The union makes ownership exclusive at compile time. Controlled read-only consumers may omit
 * `onDocumentChange`; editable controlled consumers should apply its snapshots to `document`.
 */
export type TegoSheetProps = TegoSheetSharedProps & TegoSheetOwnership;

/**
 * Imperative API exposed through a React ref while `TegoSheet` is mounted.
 *
 * @remarks
 * The handle remains stable for a mount. Calling it after unmount, or before its runtime is ready,
 * throws a `TegoSheetException`. Mutations use the same callbacks as user actions.
 *
 * @example
 * ```tsx
 * function Editor() {
 *   const sheetRef = useRef<TegoSheetHandle>(null);
 *
 *   function addBudgetSheet() {
 *     const sheet = sheetRef.current?.addSheet('Budget');
 *     if (sheet) sheetRef.current?.setCellText({ sheet, row: 0, column: 0 }, '1250');
 *   }
 *
 *   return (
 *     <>
 *       <button onClick={addBudgetSheet}>Add budget</button>
 *       <TegoSheet ref={sheetRef} defaultDocument={createSpreadsheetDocument()} />
 *     </>
 *   );
 * }
 * ```
 */
export interface TegoSheetHandle {
  /** Moves DOM focus to the spreadsheet root. */
  focus(): void;
  /** Returns an isolated snapshot of the current complete workbook. */
  getDocument(): SpreadsheetDocument;
  /** Returns cell data at an address, or `null` when the sparse cell is empty. */
  getCell(address: CellAddress): CellData | null;
  /** Returns the effective style at an address, including inherited defaults. */
  getCellStyle(address: CellAddress): CellStyle;
  /** Commits cell text at an address with change source `ref`. */
  setCellText(address: CellAddress, text: string): void;
  /** Adds a worksheet and returns its generated identifier. */
  addSheet(name?: string): SheetId;
  /** Deletes an identified worksheet and selects a replacement when necessary. */
  deleteSheet(sheet: SheetId): void;
  /** Changes the display name of an identified worksheet. */
  renameSheet(sheet: SheetId, name: string): void;
  /** Activates an identified worksheet and emits an active-sheet event with source `ref`. */
  activateSheet(sheet: SheetId): void;
  /** Activates a saved filter view as session-only presentation state. */
  activateFilterView(sheet: SheetId, viewId: string): void;
  /** Clears the active session-only filter view for a worksheet. */
  deactivateFilterView(sheet: SheetId): void;
  /** Commits the previous history state when undo is available. */
  undo(): void;
  /** Reapplies the next history state when redo is available. */
  redo(): void;
  /** Validates every configured cell rule and returns a {@link ValidationResult} with all issues. */
  validate(): ValidationResult;
  /** Recomputes canvas layout after an external container-size change. */
  recalculateLayout(): void;
}
