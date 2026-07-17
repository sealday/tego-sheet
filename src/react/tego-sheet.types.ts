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
  WorkbookData,
  WorkbookInput,
} from '../core';
import type { SheetTabsRenderer, ToolbarRenderer } from '../ui/slot-types';

export type { SheetTabsRenderer, ToolbarRenderer } from '../ui/slot-types';

/** Commit, selection, and error notifications emitted by a mounted spreadsheet. */
export interface TegoSheetCallbacks {
  /**
   * Runs after a mutation commits, before any edit, paste, or resulting selection callback.
   * `value` is the complete workbook snapshot and `change` is its {@link WorkbookChange} metadata.
   * External controlled-value replacements do not emit this callback.
   */
  readonly onChange?: (value: WorkbookData, change: WorkbookChange) => void;
  /** Runs after a worksheet is activated and reports its identifier, index, and source. */
  readonly onActiveSheetChange?: (event: ActiveSheetChangeEvent) => void;
  /** Runs after the active selection changes, including selection changes caused by a commit. */
  readonly onSelectionChange?: (selection: Selection) => void;
  /** Runs after `onChange` when a cell text edit commits. */
  readonly onCellEdit?: (event: CellEditEvent) => void;
  /** Runs after `onChange` when an internal or external paste commits. */
  readonly onPaste?: (event: PasteEvent) => void;
  /** Runs when the component handles an operation failure and exposes its structured payload. */
  readonly onError?: (error: TegoSheetError) => void;
}

/**
 * Props for the `TegoSheet` React component.
 *
 * @remarks
 * Choose controlled `value` or uncontrolled `defaultValue` when mounting. Supplying both, or
 * switching a mounted instance between those ownership modes, throws a `TegoSheetException`.
 */
export interface TegoSheetProps extends TegoSheetCallbacks {
  /**
   * Controlled workbook input owned by the parent.
   * Apply `onChange` snapshots to this prop to accept user and imperative mutations.
   */
  readonly value?: WorkbookInput;
  /** Initial workbook input for an uncontrolled component that owns subsequent mutations. */
  readonly defaultValue?: WorkbookInput;
  /** Zero-based worksheet index selected on mount. */
  readonly initialActiveSheetIndex?: number;
  /** Disables workbook mutations while preserving navigation, selection, copy, and printing. */
  readonly readOnly?: boolean;
  /** Per-instance locale identifier and message dictionary for built-in chrome. */
  readonly locale?: LocaleDefinition;
  /** Per-instance worksheet behavior and layout settings. */
  readonly options?: SheetOptions;
  /** Uses the default toolbar, hides it, or replaces it with a custom renderer. */
  readonly toolbar?: 'default' | false | ToolbarRenderer;
  /** Uses the default sheet tabs, hides them, or replaces them with a custom renderer. */
  readonly sheetTabs?: 'default' | false | SheetTabsRenderer;
  /** Additional class name appended to the root spreadsheet element. */
  readonly className?: string;
  /** Inline styles applied to the root spreadsheet element. */
  readonly style?: CSSProperties;
}

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
 *       <TegoSheet ref={sheetRef} defaultValue={[]} />
 *     </>
 *   );
 * }
 * ```
 */
export interface TegoSheetHandle {
  /** Moves DOM focus to the spreadsheet root. */
  focus(): void;
  /** Returns an isolated snapshot of the current complete workbook. */
  getValue(): WorkbookData;
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
  /** Commits the previous history state when undo is available. */
  undo(): void;
  /** Reapplies the next history state when redo is available. */
  redo(): void;
  /** Validates every configured cell rule and returns a {@link ValidationResult} with all issues. */
  validate(): ValidationResult;
  /** Prints the active worksheet as A4 portrait and reports handled print failures through `onError`. */
  print(): void;
  /** Recomputes canvas layout after an external container-size change. */
  recalculateLayout(): void;
}
