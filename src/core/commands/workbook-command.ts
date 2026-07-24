import type { CellAddress, Selection, SheetId } from '../types/coordinates';
import type { BorderMode } from '../types/options';
import type { BorderLine, CellStyle } from '../types/workbook';
import type { FilterDefinition } from '../types/options';
import type { ValidationRule } from '../types/validation';
import type { PasteMode } from '../operations/clipboard';
import type {
  CellInput,
  ConditionalFormat,
  FilterView,
  SheetObject,
} from '../../document/model/document';
import type { GroupId } from '../../document/model/ids';

/** Sets the user-facing text of one cell. */
export interface SetCellTextCommand {
  /** Command discriminator. */
  readonly type: 'set-cell-text';
  /** Cell receiving the text. */
  readonly address: CellAddress;
  /** Text to parse and store. */
  readonly text: string;
}

/** Sets one normalized typed cell input without parsing it through display text. */
export interface SetCellInputCommand {
  /** Command discriminator. */
  readonly type: 'set-cell-input';
  /** Cell receiving the typed input. */
  readonly address: CellAddress;
  /** Complete safe Workbook 2.0 input value. */
  readonly input: CellInput;
}

/** Clears cell contents within a selection. */
export interface ClearContentsCommand {
  /** Command discriminator. */
  readonly type: 'clear-contents';
  /** Cells whose contents are cleared. */
  readonly selection: Selection;
}

/** Sets editable or printable metadata on selected cells. */
export interface SetCellMetadataCommand {
  /** Command discriminator. */
  readonly type: 'set-cell-metadata';
  /** Cells receiving the metadata value. */
  readonly selection: Selection;
  /** Metadata property being changed. */
  readonly property: 'editable' | 'printable';
  /** New metadata value. */
  readonly value: boolean;
}

/** Applies a partial cell style to a selection. */
export interface SetStyleCommand {
  /** Command discriminator. */
  readonly type: 'set-style';
  /** Cells receiving the style patch. */
  readonly selection: Selection;
  /** Partial style properties to apply. */
  readonly patch: Readonly<Partial<CellStyle>>;
}

/** Applies or removes selection borders. */
export interface SetBorderCommand {
  /** Command discriminator. */
  readonly type: 'set-border';
  /** Cells receiving the border operation. */
  readonly selection: Selection;
  /** Edges affected by the border operation. */
  readonly mode: BorderMode;
  /** Border line to apply, or omission to remove it. */
  readonly line?: BorderLine;
}

/** Removes direct formatting from a selection. */
export interface ClearFormatCommand {
  /** Command discriminator. */
  readonly type: 'clear-format';
  /** Cells whose direct formatting is removed. */
  readonly selection: Selection;
}

/** Copies formatting between selections. */
export interface PaintFormatCommand {
  /** Command discriminator. */
  readonly type: 'paint-format';
  /** Cells providing the formatting. */
  readonly source: Selection;
  /** Cells receiving the formatting. */
  readonly target: Selection;
}

/** Inserts or deletes contiguous worksheet rows or columns. */
export interface IndexedSheetCommand {
  /** Structural operation discriminator. */
  readonly type: 'insert-row' | 'delete-row' | 'insert-column' | 'delete-column';
  /** Worksheet receiving the structural change. */
  readonly sheet: SheetId;
  /** Zero-based first row or column index. */
  readonly index: number;
  /** Positive number of rows or columns, defaulting to one. */
  readonly count?: number;
}

/** Sets the height of contiguous worksheet rows. */
export interface ResizeRowCommand {
  /** Command discriminator. */
  readonly type: 'set-row-height';
  /** Worksheet containing the rows. */
  readonly sheet: SheetId;
  /** Zero-based first row index. */
  readonly row: number;
  /** Row height in CSS pixels. */
  readonly height: number;
  /** Positive number of rows, defaulting to one. */
  readonly count?: number;
}

/** Shows or hides contiguous worksheet rows. */
export interface HideRowCommand {
  /** Command discriminator. */
  readonly type: 'set-row-hidden';
  /** Worksheet containing the rows. */
  readonly sheet: SheetId;
  /** Zero-based first row index. */
  readonly row: number;
  /** Whether the rows are hidden. */
  readonly hidden: boolean;
  /** Positive number of rows, defaulting to one. */
  readonly count?: number;
}

/** Sets the width of contiguous worksheet columns. */
export interface ResizeColumnCommand {
  /** Command discriminator. */
  readonly type: 'set-column-width';
  /** Worksheet containing the columns. */
  readonly sheet: SheetId;
  /** Zero-based first column index. */
  readonly column: number;
  /** Column width in CSS pixels. */
  readonly width: number;
  /** Positive number of columns, defaulting to one. */
  readonly count?: number;
}

/** Shows or hides contiguous worksheet columns. */
export interface HideColumnCommand {
  /** Command discriminator. */
  readonly type: 'set-column-hidden';
  /** Worksheet containing the columns. */
  readonly sheet: SheetId;
  /** Zero-based first column index. */
  readonly column: number;
  /** Whether the columns are hidden. */
  readonly hidden: boolean;
  /** Positive number of columns, defaulting to one. */
  readonly count?: number;
}

/** Merges or unmerges cells in a selection. */
export interface MergeCommand {
  /** Merge operation discriminator. */
  readonly type: 'merge' | 'unmerge';
  /** Cells receiving the merge operation. */
  readonly selection: Selection;
}

/** Sets frozen leading rows and columns. */
export interface SetFreezeCommand {
  /** Command discriminator. */
  readonly type: 'set-freeze';
  /** Worksheet receiving the freeze boundary. */
  readonly sheet: SheetId;
  /** Number of leading frozen rows. */
  readonly row: number;
  /** Number of leading frozen columns. */
  readonly column: number;
}

/** Adds a worksheet. */
export interface AddSheetCommand {
  /** Command discriminator. */
  readonly type: 'add-sheet';
  /** Optional requested worksheet name. */
  readonly name?: string;
}

/** Deletes a worksheet. */
export interface DeleteSheetCommand {
  /** Command discriminator. */
  readonly type: 'delete-sheet';
  /** Worksheet being deleted. */
  readonly sheet: SheetId;
}

/** Renames a worksheet. */
export interface RenameSheetCommand {
  /** Command discriminator. */
  readonly type: 'rename-sheet';
  /** Worksheet being renamed. */
  readonly sheet: SheetId;
  /** New worksheet name. */
  readonly name: string;
}

/** Pastes a selection copied within the controller. */
export interface PasteInternalCommand {
  /** Command discriminator. */
  readonly type: 'paste-internal';
  /** Original copied or cut selection. */
  readonly source: Selection;
  /** Destination selection. */
  readonly target: Selection;
  /** Content categories copied to the destination. */
  readonly mode: PasteMode;
  /** Whether the source is cleared after pasting. */
  readonly cut: boolean;
}

/** Pastes an external rectangular text matrix. */
export interface PasteExternalCommand {
  /** Command discriminator. */
  readonly type: 'paste-external';
  /** Top-left destination selection. */
  readonly target: Selection;
  /** Rectangular text matrix to parse and write. */
  readonly values: readonly (readonly string[])[];
}

/** Extends a source pattern into a target selection. */
export interface AutofillCommand {
  /** Command discriminator. */
  readonly type: 'autofill';
  /** Cells providing the fill pattern. */
  readonly source: Selection;
  /** Cells receiving the fill pattern. */
  readonly target: Selection;
  /** Content categories copied into the target. */
  readonly mode: PasteMode;
}

/** Creates one persistent row or column outline group. */
export interface GroupCommand {
  /** Command discriminator. */
  readonly type: 'group';
  /** Worksheet receiving the group. */
  readonly sheet: SheetId;
  /** Stable group definition; nesting level is derived by the controller. */
  readonly group: {
    readonly id: GroupId;
    readonly axis: 'row' | 'column';
    readonly start: number;
    readonly end: number;
    readonly collapsed: boolean;
  };
}

/** Removes one persistent outline group by stable ID. */
export interface UngroupCommand {
  /** Command discriminator. */
  readonly type: 'ungroup';
  /** Worksheet owning the group. */
  readonly sheet: SheetId;
  /** Stable group ID to remove. */
  readonly id: GroupId;
}

/** Toggles one persistent outline group's collapsed state. */
export interface ToggleGroupCommand {
  /** Command discriminator. */
  readonly type: 'toggle-group';
  /** Worksheet owning the group. */
  readonly sheet: SheetId;
  /** Stable group ID to toggle. */
  readonly id: GroupId;
}

/** Creates or replaces a worksheet filter. */
export interface SetFilterCommand {
  /** Command discriminator. */
  readonly type: 'set-filter';
  /** Filter range. */
  readonly selection: Selection;
  /** Filter definition applied to the range. */
  readonly filter: FilterDefinition;
}

/** Clears the active filter from a worksheet. */
export interface ClearFilterCommand {
  /** Command discriminator. */
  readonly type: 'clear-filter';
  /** Worksheet whose filter is cleared. */
  readonly sheet: SheetId;
}

/** Sorts the active filter range by one worksheet column. */
export interface SortCommand {
  /** Command discriminator. */
  readonly type: 'sort';
  /** Worksheet containing the filter range. */
  readonly sheet: SheetId;
  /** Absolute zero-based worksheet column. */
  readonly column: number;
  /** Sort direction. */
  readonly order: 'asc' | 'desc';
}

/** Applies a validation rule to a selection. */
export interface SetValidationCommand {
  /** Command discriminator. */
  readonly type: 'set-validation';
  /** Cells receiving the validation rule. */
  readonly selection: Selection;
  /** Validation rule to apply. */
  readonly rule: ValidationRule;
}

/** Removes validation rules from a selection. */
export interface RemoveValidationCommand {
  /** Command discriminator. */
  readonly type: 'remove-validation';
  /** Cells whose validation rules are removed. */
  readonly selection: Selection;
}

/** Creates or replaces one ordered worksheet conditional-format rule. */
export interface SetConditionalFormatCommand {
  /** Command discriminator. */
  readonly type: 'set-conditional-format';
  /** Worksheet owning the rule. */
  readonly sheet: SheetId;
  /** Zero-based rule position; the current length appends a rule. */
  readonly index: number;
  /** Complete persistent conditional-format definition. */
  readonly format: ConditionalFormat;
}

/** Removes one ordered worksheet conditional-format rule. */
export interface RemoveConditionalFormatCommand {
  /** Command discriminator. */
  readonly type: 'remove-conditional-format';
  /** Worksheet owning the rule. */
  readonly sheet: SheetId;
  /** Zero-based rule position to remove. */
  readonly index: number;
}

/** Creates or replaces one persistent worksheet filter view. */
export interface SetFilterViewCommand {
  /** Command discriminator. */
  readonly type: 'set-filter-view';
  /** Worksheet owning the view. */
  readonly sheet: SheetId;
  /** Complete persistent view definition. */
  readonly view: FilterView;
}

/** Removes one persistent worksheet filter view. */
export interface RemoveFilterViewCommand {
  /** Command discriminator. */
  readonly type: 'remove-filter-view';
  /** Worksheet owning the view. */
  readonly sheet: SheetId;
  /** Stable identifier of the view to remove. */
  readonly viewId: string;
}

/** Creates or replaces one persistent floating worksheet object. */
export interface SetSheetObjectCommand {
  /** Command discriminator. */
  readonly type: 'set-sheet-object';
  /** Worksheet owning the object. */
  readonly sheet: SheetId;
  /** Complete persistent object definition. */
  readonly object: SheetObject;
}

/** Removes one persistent floating worksheet object. */
export interface RemoveSheetObjectCommand {
  /** Command discriminator. */
  readonly type: 'remove-sheet-object';
  /** Worksheet owning the object. */
  readonly sheet: SheetId;
  /** Stable identifier of the object to remove. */
  readonly objectId: string;
}

/** Traverses one step backward through controller history. */
export interface UndoCommand {
  /** Command discriminator. */
  readonly type: 'undo';
}

/** Traverses one step forward through controller history. */
export interface RedoCommand {
  /** Command discriminator. */
  readonly type: 'redo';
}

/** Commands reserved for controller history methods and forbidden inside transactions. */
export type HistoryCommand = UndoCommand | RedoCommand;

/**
 * The closed document command contract. Task-specific operation modules handle
 * each variant; adding a variant here without an apply branch must fail rather
 * than report a successful mutation.
 */
export type WorkbookCommand =
  | SetCellTextCommand
  | SetCellInputCommand
  | ClearContentsCommand
  | SetCellMetadataCommand
  | SetStyleCommand
  | SetBorderCommand
  | ClearFormatCommand
  | PaintFormatCommand
  | IndexedSheetCommand
  | ResizeRowCommand
  | HideRowCommand
  | ResizeColumnCommand
  | HideColumnCommand
  | MergeCommand
  | SetFreezeCommand
  | AddSheetCommand
  | DeleteSheetCommand
  | RenameSheetCommand
  | PasteInternalCommand
  | PasteExternalCommand
  | AutofillCommand
  | GroupCommand
  | UngroupCommand
  | ToggleGroupCommand
  | SetFilterCommand
  | ClearFilterCommand
  | SortCommand
  | SetValidationCommand
  | RemoveValidationCommand
  | SetConditionalFormatCommand
  | RemoveConditionalFormatCommand
  | SetFilterViewCommand
  | RemoveFilterViewCommand
  | SetSheetObjectCommand
  | RemoveSheetObjectCommand
  | UndoCommand
  | RedoCommand;

export type CommandResult<Command extends WorkbookCommand> = Command extends AddSheetCommand
  ? SheetId
  : Command extends PasteInternalCommand | PasteExternalCommand
    ? readonly (readonly string[])[]
    : void;
