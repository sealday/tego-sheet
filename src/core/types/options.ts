import type { Selection, SheetId } from './coordinates';
import type { BorderLine, CellStyle } from './workbook';
import type { ValidationRule } from './validation';

/** Initial row model and default row layout. */
export interface SheetRowOptions {
  /** Row count used when an input sheet omits `rows.len`. */
  readonly initialCount?: number;
  /** Default row height in CSS pixels. */
  readonly defaultHeight?: number;
}

/** Initial column model and default column layout. */
export interface SheetColumnOptions {
  /** Column count used when an input sheet omits `cols.len`. */
  readonly initialCount?: number;
  /** Default column width in CSS pixels. */
  readonly defaultWidth?: number;
  /** Smallest width allowed during interactive column resizing. */
  readonly minimumWidth?: number;
}

/**
 * Per-instance behavior and layout options for a worksheet component.
 *
 * @remarks
 * Row, column, header, default-style, and focus settings are captured when the component mounts.
 * Remount the component to apply changes to those settings.
 */
export interface SheetOptions {
  /** Whether grid lines are visible. */
  readonly showGrid?: boolean;
  /** Whether the built-in context menu is available. */
  readonly showContextMenu?: boolean;
  /** Initial row model and default row height. */
  readonly rows?: SheetRowOptions;
  /** Initial column model and resizing limits. */
  readonly columns?: SheetColumnOptions;
  /** Width of the row-number header in CSS pixels. */
  readonly rowHeaderWidth?: number;
  /** Style inherited by cells without a more specific style. */
  readonly defaultStyle?: CellStyle;
  /** Whether the component receives focus after mounting. */
  readonly autoFocus?: boolean;
}

/** Column value filter executed by a toolbar slot action. */
export interface FilterDefinition {
  /** Zero-based column index in the worksheet. */
  readonly column: number;
  /** Whether all values or only listed values remain visible. */
  readonly operator: 'all' | 'in';
  /** Included cell text values when `operator` is `in`. */
  readonly value: readonly string[];
}

/** Region of a selection affected by a border action. */
export type BorderMode =
  | 'none'
  | 'all'
  | 'inside'
  | 'outside'
  | 'horizontal'
  | 'vertical'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right';

/** Command that a custom toolbar can request through `execute`. */
export type ToolbarAction =
  | {
      /** History or print command. */
      readonly type: 'undo' | 'redo' | 'print';
    }
  | {
      /** Format-paint or format-clearing command. */
      readonly type: 'paint-format' | 'clear-format';
    }
  | {
      /** Applies a partial style to the current selection. */
      readonly type: 'set-style';
      /** Style fields to apply without clearing unspecified fields. */
      readonly patch: Readonly<Partial<CellStyle>>;
    }
  | {
      /** Applies borders to the current selection. */
      readonly type: 'set-border';
      /** Part of the selection to receive the border. */
      readonly mode: BorderMode;
      /** Border style and color; omission uses the default border line. */
      readonly line?: BorderLine;
    }
  | {
      /** Merges the selection or separates its merged cells. */
      readonly type: 'merge' | 'unmerge';
    }
  | {
      /** Freezes panes at the active cell or removes frozen panes. */
      readonly type: 'freeze' | 'unfreeze';
    }
  | {
      /** Inserts, deletes, hides, or unhides the selected rows. */
      readonly type: 'insert-row' | 'delete-row' | 'hide-row' | 'unhide-row';
    }
  | {
      /** Inserts, deletes, hides, or unhides the selected columns. */
      readonly type: 'insert-column' | 'delete-column' | 'hide-column' | 'unhide-column';
    }
  | {
      /** Applies a validation rule to the current selection. */
      readonly type: 'set-validation';
      /** Rule to store for the selected cells. */
      readonly rule: ValidationRule;
    }
  | {
      /** Removes validation rules from the current selection. */
      readonly type: 'remove-validation';
    }
  | {
      /** Creates or updates a value filter for the selected range. */
      readonly type: 'set-filter';
      /** Column filter to apply. */
      readonly filter: FilterDefinition;
    }
  | {
      /** Removes the active worksheet filter. */
      readonly type: 'clear-filter';
    }
  | {
      /** Sorts filtered rows by the active column. */
      readonly type: 'sort';
      /** Sort direction. */
      readonly order: 'asc' | 'desc';
    };

/** State and actions supplied to a custom toolbar renderer. */
export interface ToolbarRenderProps {
  /** Current selection, or `null` when no cell is selected. */
  readonly selection: Selection | null;
  /** Effective style at the active cell. */
  readonly activeStyle: CellStyle;
  /** Whether workbook mutations are currently disabled. */
  readonly readOnly: boolean;
  /** Whether an undo history entry is available. */
  readonly canUndo: boolean;
  /** Whether a redo history entry is available. */
  readonly canRedo: boolean;
  /** Whether the current selection intersects a merged range. */
  readonly merged: boolean;
  /** Whether the active worksheet has frozen rows or columns. */
  readonly frozen: boolean;
  /** Action types unavailable for the current selection and component state. */
  readonly disabledActions: ReadonlySet<ToolbarAction['type']>;
  /** Executes an enabled toolbar action against the current component state. */
  readonly execute: (action: ToolbarAction) => void;
}

/** Immutable worksheet metadata supplied to a custom sheet-tab renderer. */
export interface SheetTabItem {
  /** Worksheet identifier used by tab actions. */
  readonly id: SheetId;
  /** Zero-based worksheet position in the workbook. */
  readonly index: number;
  /** Display name shown for the worksheet. */
  readonly name: string;
}

/** State and actions supplied to a custom sheet-tab renderer. */
export interface SheetTabsRenderProps {
  /** Ordered worksheets to display. */
  readonly sheets: readonly SheetTabItem[];
  /** Active worksheet, or `null` for an empty workbook. */
  readonly activeSheet: SheetId | null;
  /** Whether sheet-structure mutations are currently disabled. */
  readonly readOnly: boolean;
  /** Adds a worksheet and optionally assigns its display name. */
  readonly add: (name?: string) => void;
  /** Deletes the identified worksheet. */
  readonly delete: (sheet: SheetId) => void;
  /** Renames the identified worksheet. */
  readonly rename: (sheet: SheetId, name: string) => void;
  /** Makes the identified worksheet active. */
  readonly activate: (sheet: SheetId) => void;
}
