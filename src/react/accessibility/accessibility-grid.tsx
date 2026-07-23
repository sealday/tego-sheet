import { useId, useLayoutEffect, useRef } from 'react';
import type { CellPoint } from '../../document';
import type { CellPresentation } from '../../presentation';

/** Inclusive visible row and column bounds. */
export interface AccessibilityViewport {
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly columnStart: number;
  readonly columnEnd: number;
}

/** Inclusive selected cell range. */
export interface AccessibilitySelection {
  readonly start: CellPoint;
  readonly end: CellPoint;
}

/** Props for the viewport-bounded semantic spreadsheet grid. */
export interface AccessibilityGridProps {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly viewport: AccessibilityViewport;
  readonly overscan?: number;
  readonly activeCell: CellPoint;
  readonly selection: AccessibilitySelection;
  readonly resolvePresentation: (point: CellPoint) => CellPresentation;
  readonly readOnly?: boolean;
  readonly editorOpen?: boolean;
  readonly idPrefix?: string;
  readonly restoreFocus?: boolean;
  readonly onActivate?: (point: CellPoint) => void;
  readonly onRequestEdit?: (point: CellPoint) => void;
}

function bounds(count: number, start: number, end: number, overscan: number): readonly number[] {
  const first = Math.max(0, start - overscan);
  const last = Math.min(count - 1, end + overscan);
  return last < first ? [] : Array.from({ length: last - first + 1 }, (_, index) => first + index);
}

function selected(point: CellPoint, selection: AccessibilitySelection): boolean {
  const rowStart = Math.min(selection.start.row, selection.end.row);
  const rowEnd = Math.max(selection.start.row, selection.end.row);
  const columnStart = Math.min(selection.start.column, selection.end.column);
  const columnEnd = Math.max(selection.start.column, selection.end.column);
  return (
    point.row >= rowStart &&
    point.row <= rowEnd &&
    point.column >= columnStart &&
    point.column <= columnEnd
  );
}

function selectionSize(selection: AccessibilitySelection): number {
  return (
    (Math.abs(selection.end.row - selection.start.row) + 1) *
    (Math.abs(selection.end.column - selection.start.column) + 1)
  );
}

/** Virtual semantic grid mirroring shared presentation without duplicating sheet DOM. */
export function AccessibilityGrid({
  rowCount,
  columnCount,
  viewport,
  overscan = 1,
  activeCell,
  selection,
  resolvePresentation,
  readOnly = false,
  editorOpen = false,
  idPrefix,
  restoreFocus = true,
  onActivate,
  onRequestEdit,
}: AccessibilityGridProps) {
  const activeReference = useRef<HTMLDivElement>(null);
  const previousEditorOpen = useRef(editorOpen);
  useLayoutEffect(() => {
    if (restoreFocus && previousEditorOpen.current && !editorOpen) activeReference.current?.focus();
    previousEditorOpen.current = editorOpen;
  }, [editorOpen, restoreFocus]);
  const rows = bounds(rowCount, viewport.rowStart, viewport.rowEnd, overscan);
  const columns = bounds(columnCount, viewport.columnStart, viewport.columnEnd, overscan);
  const selectedCount = selectionSize(selection);
  const selectionDescriptionId = useId();
  const generatedCellIdPrefix = useId();
  const cellIdPrefix = idPrefix ?? generatedCellIdPrefix;

  return (
    <div
      role="grid"
      aria-rowcount={rowCount}
      aria-colcount={columnCount}
      aria-multiselectable={selectedCount > 1}
      aria-describedby={selectionDescriptionId}
    >
      <div id={selectionDescriptionId} aria-live="polite" hidden={selectedCount <= 1}>
        {selectedCount <= 1 ? '' : `Selected ${selectedCount} cells`}
      </div>
      {rows.map((row) => (
        <div role="row" aria-rowindex={row + 1} key={row}>
          {columns.map((column) => {
            const point = { row, column };
            const presentation = resolvePresentation(point);
            if (presentation.visibility.hidden) return null;
            const active = row === activeCell.row && column === activeCell.column;
            return (
              <div
                id={`${cellIdPrefix}-r${row}-c${column}`}
                role="gridcell"
                aria-rowindex={row + 1}
                aria-colindex={column + 1}
                aria-label={presentation.accessibility.label}
                aria-readonly={readOnly || presentation.accessibility.readOnly}
                aria-invalid={presentation.accessibility.invalid || undefined}
                aria-checked={
                  presentation.accessibility.role === 'checkbox'
                    ? presentation.accessibility.checked
                    : undefined
                }
                aria-haspopup={
                  presentation.accessibility.role === 'combobox' ? 'listbox' : undefined
                }
                aria-selected={selected(point, selection)}
                key={column}
                ref={active ? activeReference : undefined}
                tabIndex={active ? 0 : -1}
                onClick={() => onActivate?.(point)}
                onDoubleClick={() => onRequestEdit?.(point)}
              >
                {presentation.formattedText}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}
