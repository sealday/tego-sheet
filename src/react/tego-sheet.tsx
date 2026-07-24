import {
  forwardRef,
  useCallback,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useReducer,
  useRef,
  useState,
  type ForwardedRef,
} from 'react';
import {
  parseA1,
  parseA1Range,
  rangesIntersect,
  selectCellStyle,
  TegoSheetException,
  type ChangeSource,
  type FilterDefinition,
  type Selection,
  type SheetId,
  type SheetTabsRenderProps,
  type ToolbarAction,
  type ToolbarRenderProps,
  type ValidationRule,
} from '../core';
import type { WorkbookCommand } from '../core/commands/workbook-command';
import { createEventDispatcher } from './adapters/event-dispatcher';
import { deletionSplitsMerge } from '../core/operations/structure';
import { createEngineAdapterSlot, useCanvasEngine } from './hooks/use-canvas-engine';
import { useControllerEpoch, type ControllerEpoch } from './hooks/use-controller-epoch';
import { useInteractionManager } from './hooks/use-interaction-manager';
import { createSelectionState, visibleCellRange, type InteractionManager } from '../engine';
import { useSheetChromeState } from './hooks/use-sheet-chrome-state';
import { useCellEditorRuntime, type ActiveCellEditor } from './hooks/use-cell-editor-runtime';
import {
  useTegoSheetHandle,
  type TegoSheetHandleRuntime,
  type TegoSheetRuntimeAuthority,
} from './hooks/use-tego-sheet-handle';
import {
  useMountOptionWarnings,
  type TegoSheetMountOptions,
} from './hooks/use-mount-option-warnings';
import {
  useControlledWorkbook,
  type ControlledWorkbookRuntime,
} from './hooks/use-controlled-workbook';
import type { TegoSheetCallbacks, TegoSheetHandle, TegoSheetProps } from './tego-sheet.types';
import { EmptyWorkbook } from '../ui/empty-workbook';
import { SheetChrome } from '../ui/sheet-chrome';
import type { ContextMenuAction } from '../ui/menus/context-menu';
import { createTranslator } from '../ui/translate';
import {
  activeSheetData,
  filterCommandSelection,
  filterValuesForSelection,
} from './sheet-chrome-runtime';
import { AccessibilityGrid } from './accessibility/accessibility-grid';
import { AccessibilityObjects } from './accessibility/accessibility-objects';
import { createPresentationCache, createPresentationResolver } from '../presentation';
import type {
  ValidationEngineOptions,
  ValidationResult as AdvancedValidationResult,
} from '../validation';
import { createPresentationValidationResolver } from './adapters/presentation-adapter';
import { compileSpreadsheetTemplate, renderSpreadsheetTemplate } from '../template';
import { TemplateDesigner } from './template-designer';
import { TemplatePreview } from './preview';
import { applyDocumentFilterView } from '../views';
import { projectSheetObjectsToViewport } from './adapters/object-adapter';

function callbacksFromProps(props: TegoSheetProps): TegoSheetCallbacks {
  return {
    onActiveSheetChange: props.onActiveSheetChange,
    onCellEdit: props.onCellEdit,
    onDocumentChange: props.onDocumentChange,
    onError: props.onError,
    onPaste: props.onPaste,
    onSelectionChange: props.onSelectionChange,
  };
}

interface CallbackStore {
  readonly get: () => TegoSheetCallbacks;
  readonly set: (callbacks: TegoSheetCallbacks) => void;
}

function createCallbackStore(initial: TegoSheetCallbacks): CallbackStore {
  let current = initial;
  return {
    get: () => current,
    set(callbacks) {
      current = callbacks;
    },
  };
}

function clonePublic<T>(value: T): T {
  if (Array.isArray(value)) return value.map(clonePublic) as T;
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, clonePublic(item)]),
    ) as T;
  }
  return value;
}

function classNames(value: string | undefined): string {
  return value === undefined || value.trim().length === 0 ? 'tego-sheet' : `tego-sheet ${value}`;
}

function contractViolation(message: string): TegoSheetException {
  return new TegoSheetException({
    code: 'INVALID_COMMAND',
    message,
    recoverable: false,
  });
}

const MUTATING_ACTIONS = new Set<ToolbarAction['type']>([
  'undo',
  'redo',
  'paint-format',
  'clear-format',
  'set-style',
  'set-border',
  'merge',
  'unmerge',
  'freeze',
  'unfreeze',
  'insert-row',
  'delete-row',
  'hide-row',
  'unhide-row',
  'insert-column',
  'delete-column',
  'hide-column',
  'unhide-column',
  'set-validation',
  'remove-validation',
  'set-filter',
  'clear-filter',
  'sort',
]);

const SELECTION_ACTIONS = new Set<ToolbarAction['type']>([
  'paint-format',
  'clear-format',
  'set-style',
  'set-border',
  'merge',
  'unmerge',
  'freeze',
  'insert-row',
  'delete-row',
  'hide-row',
  'unhide-row',
  'insert-column',
  'delete-column',
  'hide-column',
  'unhide-column',
  'set-validation',
  'remove-validation',
  'set-filter',
  'sort',
]);

interface SlotRuntime extends TegoSheetHandleRuntime {
  readonly selection: Selection | null;
  readonly readOnly: boolean;
  readonly validation: ValidationEngineOptions;
  readonly confirmValidationWarning?: (
    result: AdvancedValidationResult,
  ) => boolean | Promise<boolean>;
}

interface DialogAuthority {
  readonly source: 'toolbar' | 'context-menu';
  readonly selection: Selection | null;
}

interface FilterDialogAuthority extends DialogAuthority {
  readonly values: readonly string[];
}

type SlotRuntimeAuthority = TegoSheetRuntimeAuthority<SlotRuntime>;

function uiError(runtime: SlotRuntime, message: string): void {
  runtime.dispatcher.reportUiError({
    code: 'INVALID_COMMAND',
    message,
    recoverable: true,
  });
}

function runtimeSheet(runtime: SlotRuntime) {
  return activeSheetData(runtime.controller.getSnapshot(), runtime.activeSheet);
}

function runtimeMerges(runtime: SlotRuntime) {
  return (runtimeSheet(runtime)?.merges ?? []).map(parseA1Range);
}

function mergedSelection(runtime: SlotRuntime): boolean {
  const selection = runtime.selection;
  return (
    selection !== null &&
    runtimeMerges(runtime).some((merge) => rangesIntersect(merge, selection.range))
  );
}

function frozenSheet(runtime: SlotRuntime): boolean {
  const point = parseA1(runtimeSheet(runtime)?.freeze ?? 'A1');
  return point.row > 0 || point.column > 0;
}

function disabledToolbarActions(runtime: SlotRuntime): Set<ToolbarAction['type']> {
  const snapshot = runtime.controller.getSnapshot();
  const disabled = new Set<ToolbarAction['type']>();
  const selection = runtime.selection;
  const sheet = runtimeSheet(runtime);
  if (runtime.readOnly || snapshot.readOnly) {
    for (const action of MUTATING_ACTIONS) disabled.add(action);
  }
  if (!snapshot.canUndo) disabled.add('undo');
  if (!snapshot.canRedo) disabled.add('redo');
  if (selection === null) {
    for (const action of SELECTION_ACTIONS) disabled.add(action);
  }
  if (sheet === null) {
    disabled.add('clear-filter');
    disabled.add('sort');
    disabled.add('unfreeze');
  }
  const merged = mergedSelection(runtime);
  const singleCell =
    selection !== null &&
    selection.range.start.row === selection.range.end.row &&
    selection.range.start.column === selection.range.end.column;
  if (merged || singleCell) disabled.add('merge');
  if (!merged) disabled.add('unmerge');
  const frozen = frozenSheet(runtime);
  if (frozen) disabled.add('freeze');
  else disabled.add('unfreeze');
  if (selection?.active.row === 0 && selection.active.column === 0) disabled.add('freeze');

  let filterRange: ReturnType<typeof parseA1Range> | null = null;
  const filterReference = sheet?.autofilter?.ref;
  if (filterReference !== undefined) {
    try {
      filterRange = parseA1Range(filterReference);
    } catch {
      filterRange = null;
    }
  }
  if (filterRange === null) {
    disabled.add('clear-filter');
    disabled.add('sort');
  } else if (
    selection === null ||
    selection.active.column < filterRange.start.column ||
    selection.active.column > filterRange.end.column
  )
    disabled.add('sort');

  if (selection !== null && sheet !== null) {
    if (deletionSplitsMerge(sheet, 'row', selection.range.start.row, selection.range.end.row))
      disabled.add('delete-row');
    if (
      deletionSplitsMerge(sheet, 'column', selection.range.start.column, selection.range.end.column)
    )
      disabled.add('delete-column');
  }
  return disabled;
}

function readonlySet<Value>(source: ReadonlySet<Value>): ReadonlySet<Value> {
  const values = Array.from(source);
  return Object.freeze({
    get size() {
      return values.length;
    },
    has: (value: Value) => source.has(value),
    entries: () => values.map((value) => [value, value] as [Value, Value]).values(),
    keys: () => values.values(),
    values: () => values.values(),
    forEach(
      callback: (value: Value, key: Value, set: ReadonlySet<Value>) => void,
      thisArg?: unknown,
    ) {
      for (const value of values) callback.call(thisArg, value, value, this);
    },
    [Symbol.iterator]: () => values.values(),
  });
}

function toolbarCommand(runtime: SlotRuntime, action: ToolbarAction): WorkbookCommand | null {
  const selection = runtime.selection;
  switch (action.type) {
    case 'paint-format':
      return null;
    case 'undo':
      return { type: 'undo' };
    case 'redo':
      return { type: 'redo' };
    case 'unfreeze':
      return runtime.activeSheet === null
        ? null
        : { type: 'set-freeze', sheet: runtime.activeSheet, row: 0, column: 0 };
    case 'clear-filter':
      return runtime.activeSheet === null
        ? null
        : { type: 'clear-filter', sheet: runtime.activeSheet };
  }
  if (selection === null) return null;
  const { sheet, active, range } = selection;
  switch (action.type) {
    case 'clear-format':
      return { type: 'clear-format', selection };
    case 'set-style':
      return { type: 'set-style', selection, patch: action.patch };
    case 'set-border':
      return {
        type: 'set-border',
        selection,
        mode: action.mode,
        line: action.line,
      };
    case 'merge':
    case 'unmerge':
      return { type: action.type, selection };
    case 'freeze':
      return {
        type: 'set-freeze',
        sheet,
        row: active.row,
        column: active.column,
      };
    case 'insert-row':
    case 'delete-row':
      return {
        type: action.type,
        sheet,
        index: range.start.row,
        count: range.end.row - range.start.row + 1,
      };
    case 'insert-column':
    case 'delete-column':
      return {
        type: action.type,
        sheet,
        index: range.start.column,
        count: range.end.column - range.start.column + 1,
      };
    case 'hide-row':
    case 'unhide-row':
      return {
        type: 'set-row-hidden',
        sheet,
        row: range.start.row,
        count: range.end.row - range.start.row + 1,
        hidden: action.type === 'hide-row',
      };
    case 'hide-column':
    case 'unhide-column':
      return {
        type: 'set-column-hidden',
        sheet,
        column: range.start.column,
        count: range.end.column - range.start.column + 1,
        hidden: action.type === 'hide-column',
      };
    case 'set-validation':
      return { type: 'set-validation', selection, rule: action.rule };
    case 'remove-validation':
      return { type: 'remove-validation', selection };
    case 'set-filter': {
      const sheetData = runtimeSheet(runtime);
      return {
        type: 'set-filter',
        selection: sheetData === null ? selection : filterCommandSelection(sheetData, selection),
        filter: action.filter,
      };
    }
    case 'sort':
      return {
        type: 'sort',
        sheet,
        column: active.column,
        order: action.order,
      };
  }
}

function executeAction(
  runtime: SlotRuntime,
  action: ToolbarAction,
  source: 'toolbar' | 'context-menu',
  selection: Selection | null = runtime.selection,
): void {
  const actionRuntime =
    selection === runtime.selection
      ? runtime
      : {
          ...runtime,
          activeSheet: selection?.sheet ?? runtime.activeSheet,
          selection,
        };
  if (disabledToolbarActions(actionRuntime).has(action.type)) {
    uiError(
      runtime,
      `${source === 'toolbar' ? 'Toolbar' : 'Context-menu'} action "${action.type}" is unavailable`,
    );
    return;
  }
  const command = toolbarCommand(actionRuntime, action);
  if (command === null) {
    uiError(
      runtime,
      `${source === 'toolbar' ? 'Toolbar' : 'Context-menu'} action "${action.type}" cannot run in the current view`,
    );
    return;
  }
  runtime.dispatcher.dispatchUi(command, source);
}

function addSheetFromTabs(authority: SlotRuntimeAuthority, name?: string): void {
  const capture = authority.capture();
  const { runtime } = capture;
  if (runtime.readOnly || runtime.controller.getSnapshot().readOnly) {
    uiError(runtime, 'Sheet tabs cannot add a sheet while the workbook is read-only');
    return;
  }
  const wasEmpty = runtime.controller.getSnapshot().sheets.length === 0;
  const outcome = runtime.dispatcher.dispatchUi(
    name === undefined ? { type: 'add-sheet' } : { type: 'add-sheet', name },
    'sheet-tabs',
  );
  if (wasEmpty && outcome.status === 'committed' && typeof outcome.commit.result === 'string')
    authority.compareAndSetActiveSheet(capture, outcome.commit.result as SheetId);
}

function deleteSheetFromTabs(authority: SlotRuntimeAuthority, sheet: SheetId): void {
  const capture = authority.capture();
  const { runtime } = capture;
  if (runtime.readOnly || runtime.controller.getSnapshot().readOnly) {
    uiError(runtime, 'Sheet tabs cannot delete a sheet while the workbook is read-only');
    return;
  }
  const before = runtime.controller.getSnapshot();
  const removedIndex = before.sheets.findIndex((item) => item.id === sheet);
  const outcome = runtime.dispatcher.dispatchUi({ type: 'delete-sheet', sheet }, 'sheet-tabs');
  if (outcome.status !== 'committed' || runtime.activeSheet !== sheet) return;
  const after = runtime.controller.getSnapshot();
  const replacementIndex = Math.min(removedIndex, after.sheets.length - 1);
  authority.compareAndSetActiveSheet(
    capture,
    replacementIndex < 0 ? null : after.sheets[replacementIndex]!.id,
  );
}

function renameSheetFromTabs(runtime: SlotRuntime, sheet: SheetId, name: string): void {
  if (runtime.readOnly || runtime.controller.getSnapshot().readOnly) {
    uiError(runtime, 'Sheet tabs cannot rename a sheet while the workbook is read-only');
    return;
  }
  runtime.dispatcher.dispatchUi({ type: 'rename-sheet', sheet, name }, 'sheet-tabs');
}

function activateSheetFromTabs(
  authority: SlotRuntimeAuthority,
  runtime: SlotRuntime,
  sheet: SheetId,
): void {
  const snapshot = runtime.controller.getSnapshot();
  const index = snapshot.sheets.findIndex((item) => item.id === sheet);
  if (index < 0) {
    uiError(runtime, `Unknown sheet ID: ${sheet}`);
    return;
  }
  authority.activate(sheet);
  runtime.dispatcher.emitActiveSheetChange({
    sheet,
    index,
    source: 'sheet-tabs',
  });
}

function CommitAuthority(props: { readonly commit: () => void }) {
  const { commit } = props;
  useLayoutEffect(() => {
    commit();
  }, [commit]);
  return null;
}

type RuntimeProps = TegoSheetProps & {
  readonly controlled: ControlledWorkbookRuntime;
  readonly epoch: ControllerEpoch;
  readonly mountOptions: TegoSheetMountOptions;
  readonly mountActiveSheetIndex: number | undefined;
};

function Runtime(props: RuntimeProps, forwardedRef: ForwardedRef<TegoSheetHandle>) {
  const rootRef = useRef<HTMLDivElement>(null);
  const focusingSurfaceRef = useRef(false);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const interactionManagerRef = useRef<InteractionManager | null>(null);
  const [engineSlot] = useState(createEngineAdapterSlot);
  const [accessibilityCache] = useState(() =>
    createPresentationCache({
      maximumEntries: 2_000,
      maximumBytes: 2 * 1024 * 1024,
    }),
  );
  const [callbackStore] = useState(() => createCallbackStore(callbacksFromProps(props)));
  const initialOptions = props.mountOptions;
  const initialActiveSheetIndex = props.mountActiveSheetIndex ?? 0;
  const [initialSheetCount] = useState(() => props.epoch.snapshot.sheets.length);
  const initialWorkbookWasEmpty = initialSheetCount === 0;
  const [activeRequest, setActiveRequest] = useState(() => ({
    index: initialWorkbookWasEmpty ? 0 : initialActiveSheetIndex,
    sheet: null as SheetId | null,
  }));
  const [selection, setSelection] = useState<Selection | null>(null);
  const [selectedObject, setSelectedObject] = useState<{
    readonly sheet: SheetId;
    readonly objectId: string;
  } | null>(null);
  const [accessibilityViewportRevision, refreshAccessibilityViewport] = useReducer(
    (value: number) => value + 1,
    0,
  );
  const [filterViewRevision, refreshFilterView] = useReducer((value: number) => value + 1, 0);
  const controller = props.epoch.controller;
  const isActive = props.epoch.isActive;
  const requestSurfaceFocus = useCallback(() => {
    const root = rootRef.current;
    if (
      !isActive() ||
      root === null ||
      root.ownerDocument.activeElement === root ||
      focusingSurfaceRef.current
    )
      return;
    focusingSurfaceRef.current = true;
    try {
      root.focus({ preventScroll: true });
    } finally {
      focusingSurfaceRef.current = false;
    }
  }, [isActive]);
  const {
    editor,
    editorRef,
    contextMenu,
    filterOpen,
    validationOpen,
    notification,
    paintSource,
    replaceEditor,
    cancelTransient,
    requestContextMenu,
    closeContextMenu,
    closeFilter,
    closeValidation,
    openFilter,
    openValidation,
    setNotification,
    togglePaintSource,
    consumePaintSource,
  } = useSheetChromeState<ActiveCellEditor>(isActive, requestSurfaceFocus);
  const [engineGeneration, signalEngineReady] = useReducer((value: number) => value + 1, 0);
  const runtimeAuthority = useTegoSheetHandle<SlotRuntime>(forwardedRef);

  const sheets = props.epoch.snapshot.sheets;
  const setActiveSheet = useCallback(
    (sheet: SheetId | null) => {
      setActiveRequest((current) => {
        const index =
          sheet === null
            ? current.index
            : props.epoch.controller.getSheetIds().findIndex((candidate) => candidate === sheet);
        const nextIndex = index < 0 ? current.index : index;
        if (current.sheet === sheet && current.index === nextIndex) return current;
        return { index: nextIndex, sheet };
      });
    },
    [props.epoch.controller],
  );
  const activeSheet =
    sheets.length === 0
      ? null
      : sheets.some((sheet) => sheet.id === activeRequest.sheet)
        ? activeRequest.sheet
        : (sheets[Math.min(activeRequest.index, sheets.length - 1)]?.id ?? sheets[0]!.id);
  const clippedActiveIndex =
    activeSheet === null
      ? 0
      : Math.max(
          0,
          sheets.findIndex((sheet) => sheet.id === activeSheet),
        );

  const dispatcher = useMemo(
    () =>
      createEventDispatcher({
        controller,
        getCallbacks: callbackStore.get,
        getControlledNotificationVersion: props.controlled.getNotificationVersion,
        isActive,
        onUiError: (error) => {
          if (isActive()) setNotification(error);
        },
        recordControlledCheckpoint: props.controlled.recordCheckpoint,
        schedulePaint: () => engineSlot.get()?.render(controller.getSnapshot(), activeSheet),
      }),
    [
      activeSheet,
      callbackStore,
      controller,
      engineSlot,
      isActive,
      props.controlled.getNotificationVersion,
      props.controlled.recordCheckpoint,
      setNotification,
    ],
  );
  const reportRenderError = useCallback(
    (cause: unknown) => {
      dispatcher.reportUiError({
        code: 'RENDER_FAILED',
        message: 'Rendering the workbook failed',
        recoverable: true,
        cause,
      });
    },
    [dispatcher],
  );

  const renderRuntime: SlotRuntime = {
    activeSheet,
    controller: props.epoch.controller,
    defaultStyle: initialOptions.defaultStyle,
    dispatcher,
    engineSlot,
    isActive: props.epoch.isActive,
    readOnly: props.readOnly ?? false,
    validation: props.validationEngine ?? {},
    ...(props.confirmValidationWarning === undefined
      ? {}
      : { confirmValidationWarning: props.confirmValidationWarning }),
    root: null,
    refreshFilterView,
    selection,
    setActiveSheet,
  };
  const renderToken = {};
  const commitRuntime = () => {
    if (!props.epoch.isActive()) {
      runtimeAuthority.deactivate();
      return;
    }
    callbackStore.set(callbacksFromProps(props));
    if (props.epoch.controller.getSnapshot().readOnly !== renderRuntime.readOnly) {
      props.epoch.controller.setReadOnly(renderRuntime.readOnly);
      props.epoch.store.refresh();
    }
    engineSlot.get()?.updateReadOnly(renderRuntime.readOnly);
    runtimeAuthority.commit(renderToken, {
      ...renderRuntime,
      root: rootRef.current,
    });
  };
  const rootCallback = useCallback(
    (node: HTMLDivElement | null) => {
      rootRef.current = node;
      if (node === null) runtimeAuthority.deactivate();
      else runtimeAuthority.patchRoot(node);
    },
    [runtimeAuthority],
  );

  if (
    !initialWorkbookWasEmpty &&
    initialSheetCount > 0 &&
    (!Number.isSafeInteger(initialActiveSheetIndex) ||
      initialActiveSheetIndex < 0 ||
      initialActiveSheetIndex >= initialSheetCount)
  ) {
    throw contractViolation('initialActiveSheetIndex must refer to an initial sheet');
  }

  const { commitEditor, refreshEditorAnchor, requestEdit } = useCellEditorRuntime({
    editorRef,
    isActive,
    replaceEditor,
    runtimeAuthority,
    setSelection,
  });
  const handleViewportChange = useCallback(() => {
    refreshEditorAnchor();
    refreshAccessibilityViewport();
  }, [refreshEditorAnchor]);
  const requestDelete = useCallback(
    (target: Selection, source: ChangeSource) => {
      if (!isActive()) return;
      runtimeAuthority.require().dispatcher.dispatchUi(
        {
          type: 'clear-contents',
          selection: target,
        },
        source,
      );
    },
    [isActive, runtimeAuthority],
  );
  const requestFormat = useCallback(
    (format: 'bold' | 'italic' | 'underline') => {
      if (!isActive()) return;
      const runtime = runtimeAuthority.require();
      if (runtime.selection === null) return;
      const sheet = runtimeSheet(runtime);
      const current =
        sheet === null
          ? {}
          : selectCellStyle(
              sheet,
              runtime.selection.active.row,
              runtime.selection.active.column,
              runtime.defaultStyle,
            );
      runtime.dispatcher.dispatchUi(
        {
          type: 'set-style',
          selection: runtime.selection,
          patch:
            format === 'underline'
              ? { underline: current.underline !== true }
              : {
                  font: {
                    ...current.font,
                    [format]: current.font?.[format] !== true,
                  },
                },
        },
        'keyboard',
      );
    },
    [isActive, runtimeAuthority],
  );

  useLayoutEffect(() => {
    if (
      !props.epoch.isActive() ||
      (activeRequest.sheet === activeSheet && activeRequest.index === clippedActiveIndex)
    )
      return;
    const capturedRequest = activeRequest;
    // The controller's committed sheet set is external state; identity CAS prevents
    // this synchronization from overwriting a newer explicit active-sheet decision.
    // oxlint-disable-next-line react/react-compiler
    setActiveRequest((current) =>
      current !== capturedRequest ? current : { index: clippedActiveIndex, sheet: activeSheet },
    );
  }, [activeRequest, activeSheet, clippedActiveIndex, props.epoch]);

  // Register this layout cleanup before the canvas cleanup so browser listeners
  // are always released before the engine subscription and render scheduler.
  useInteractionManager({
    activeSheet,
    dispatcher,
    engineGeneration,
    engineSlot,
    managerRef: interactionManagerRef,
    epoch: props.epoch,
    rootRef,
    surfaceRef: canvasRef,
    showContextMenu: props.options?.showContextMenu,
    minimumColumnWidth: initialOptions.columns?.minimumWidth,
    onSelectionChange: setSelection,
    onViewportChange: handleViewportChange,
    commitEditor,
    requestCancelTransient: cancelTransient,
    requestContextMenu,
    requestDelete,
    requestEdit,
    requestFormat,
    requestSurfaceFocus,
  });
  useCanvasEngine({
    activeSheet,
    canvasRef,
    enabled: sheets.length > 0,
    filterViewRevision,
    engineSlot,
    epoch: props.epoch,
    locale: props.locale,
    onReady: signalEngineReady,
    onRenderError: reportRenderError,
    onSelectionChange: setSelection,
    rootRef,
    sheetOptions: initialOptions,
    showGrid: props.options?.showGrid,
    renderEnvironment: props.renderEnvironment,
    onObjectDiagnostics: props.onDiagnostics,
  });
  const reconciliationVersion = props.controlled.getNotificationVersion();
  const transientAuthorityRef = useRef({
    activeSheet,
    readOnly: renderRuntime.readOnly,
    reconciliationVersion,
  });
  useLayoutEffect(() => {
    const previous = transientAuthorityRef.current;
    transientAuthorityRef.current = {
      activeSheet,
      readOnly: renderRuntime.readOnly,
      reconciliationVersion,
    };
    if (
      renderRuntime.readOnly ||
      previous.activeSheet !== activeSheet ||
      previous.reconciliationVersion !== reconciliationVersion
    )
      cancelTransient();
  }, [activeSheet, cancelTransient, reconciliationVersion, renderRuntime.readOnly]);

  useLayoutEffect(() => {
    if (selection === null || !isActive()) return;
    const source = consumePaintSource(selection);
    if (source === null) return;
    runtimeAuthority.require().dispatcher.dispatchUi(
      {
        type: 'paint-format',
        source,
        target: selection,
      },
      'toolbar',
    );
  }, [consumePaintSource, isActive, runtimeAuthority, selection]);

  useLayoutEffect(() => {
    if (initialOptions.autoFocus === true) rootRef.current?.focus();
  }, [initialOptions.autoFocus]);

  const execute = (action: ToolbarAction) => {
    const runtime = runtimeAuthority.committed(renderToken);
    if (runtime === null) return;
    if (action.type !== 'paint-format') {
      executeAction(runtime, action, 'toolbar');
      return;
    }
    if (disabledToolbarActions(runtime).has('paint-format') || runtime.selection === null) {
      uiError(runtime, 'Toolbar action "paint-format" is unavailable');
      return;
    }
    togglePaintSource(clonePublic(runtime.selection));
  };
  const executeContext = (action: ContextMenuAction) => {
    const runtime = runtimeAuthority.committed(renderToken);
    if (runtime === null) return;
    const manager = interactionManagerRef.current;
    switch (action.type) {
      case 'copy':
        if (manager === null) uiError(runtime, 'Context-menu copy is unavailable');
        else void manager.copy();
        return;
      case 'cut':
        if (manager === null) uiError(runtime, 'Context-menu cut is unavailable');
        else void manager.copy(undefined, true);
        return;
      case 'paste':
      case 'paste-value':
      case 'paste-format':
        if (manager === null) uiError(runtime, 'Context-menu paste is unavailable');
        else
          void manager.paste(
            undefined,
            action.type === 'paste' ? 'all' : action.type === 'paste-value' ? 'value' : 'format',
            'context-menu',
          );
        return;
      case 'clear-contents':
        if (runtime.selection === null)
          uiError(runtime, 'Context-menu clear contents is unavailable');
        else
          runtime.dispatcher.dispatchUi(
            { type: 'clear-contents', selection: runtime.selection },
            'context-menu',
          );
        return;
      case 'set-cell-metadata':
        if (runtime.selection === null)
          uiError(runtime, 'Context-menu cell metadata is unavailable');
        else
          runtime.dispatcher.dispatchUi(
            { ...action, selection: runtime.selection },
            'context-menu',
          );
        return;
      default:
        executeAction(runtime, action, 'context-menu');
    }
  };
  const [filterAuthority, setFilterAuthority] = useState<FilterDialogAuthority | null>(null);
  const validationAuthorityRef = useRef<DialogAuthority | null>(null);
  const captureDialogAuthority = (source: DialogAuthority['source']): DialogAuthority | null => {
    const runtime = runtimeAuthority.committed(renderToken);
    return runtime === null
      ? null
      : {
          source,
          selection: runtime.selection === null ? null : clonePublic(runtime.selection),
        };
  };
  const openFilterFor = (source: DialogAuthority['source']) => {
    const authority = captureDialogAuthority(source);
    if (authority === null || authority.selection === null) return;
    const sheet = activeSheetData(props.epoch.snapshot, authority.selection.sheet);
    if (sheet === null) return;
    let values: readonly string[];
    try {
      values = filterValuesForSelection(sheet, authority.selection);
    } catch (cause) {
      dispatcher.reportUiError({
        code: 'INVALID_COMMAND',
        message: 'Filter values exceed the supported resource limit',
        recoverable: true,
        cause,
      });
      return;
    }
    setFilterAuthority({ ...authority, values });
    openFilter();
  };
  const openToolbarFilter = () => openFilterFor('toolbar');
  const openToolbarValidation = () => {
    const authority = captureDialogAuthority('toolbar');
    if (authority === null) return;
    validationAuthorityRef.current = authority;
    openValidation();
  };
  const openContextFilter = () => openFilterFor('context-menu');
  const openContextValidation = () => {
    const authority = captureDialogAuthority('context-menu');
    if (authority === null) return;
    validationAuthorityRef.current = authority;
    openValidation();
  };
  const tabActions = {
    add(name?: string) {
      if (runtimeAuthority.committed(renderToken) !== null) {
        addSheetFromTabs(runtimeAuthority, name);
      }
    },
    delete(sheet: SheetId) {
      if (runtimeAuthority.committed(renderToken) !== null) {
        deleteSheetFromTabs(runtimeAuthority, sheet);
      }
    },
    rename(sheet: SheetId, name: string) {
      const runtime = runtimeAuthority.committed(renderToken);
      if (runtime !== null) renameSheetFromTabs(runtime, sheet, name);
    },
    activate(sheet: SheetId) {
      const runtime = runtimeAuthority.committed(renderToken);
      if (runtime !== null) activateSheetFromTabs(runtimeAuthority, runtime, sheet);
    },
  };
  const addFirstSheet = () => tabActions.add();
  const activeData = runtimeSheet(renderRuntime);
  const activeDocumentSheet =
    activeSheet === null
      ? undefined
      : props.epoch.snapshot.document.workbook.sheets.find(
          ({ id }) => id === (activeSheet as string),
        );
  const activeSelectedObjectId =
    selectedObject !== null &&
    selectedObject.sheet === activeSheet &&
    activeDocumentSheet?.objects.some(({ id }) => id === selectedObject.objectId) === true
      ? selectedObject.objectId
      : null;
  useLayoutEffect(() => {
    engineSlot.get()?.setSelectedObject(activeSelectedObjectId);
  }, [activeSelectedObjectId, activeSheet, engineGeneration, engineSlot]);
  const visibleObjectProjections = (() => {
    void accessibilityViewportRevision;
    const viewport = engineSlot.get()?.interactionSnapshot()?.viewport;
    if (activeDocumentSheet === undefined || viewport === undefined) return [];
    return projectSheetObjectsToViewport(
      activeDocumentSheet.objects,
      props.epoch.snapshot.document.resources.items,
      viewport,
    );
  })();
  void filterViewRevision;
  const activeFilterView =
    activeSheet === null ? undefined : props.epoch.controller.getActiveFilterView(activeSheet);
  const accessibilityRows = useMemo(() => {
    const rowCount =
      activeData === null
        ? 0
        : Math.max(
            selection?.active.row ?? 0,
            typeof activeData.rows?.len === 'number' ? activeData.rows.len - 1 : 0,
            activeFilterView?.range.end.row ?? 0,
          ) + 1;
    if (activeFilterView === undefined) {
      return Array.from({ length: rowCount }, (_, row) => row);
    }
    const projection = applyDocumentFilterView({
      document: props.epoch.snapshot.document,
      formulaValues: new Map(
        props.epoch.snapshot.calculation.values.map(({ address, value }) => [address, value]),
      ),
      view: activeFilterView,
      locale:
        props.locale?.id ?? props.epoch.snapshot.document.workbook.settings.localeHint ?? 'en-US',
      limits: {
        maxRows: Math.max(1, activeFilterView.range.end.row - activeFilterView.range.start.row + 1),
      },
    });
    const bodyStart = activeFilterView.range.start.row + 1;
    const bodyEnd = Math.min(rowCount - 1, activeFilterView.range.end.row);
    return [
      ...Array.from({ length: Math.min(rowCount, bodyStart) }, (_, row) => row),
      ...projection.rowOrder.filter((row) => row <= bodyEnd && !projection.hiddenRows.has(row)),
      ...Array.from(
        { length: Math.max(0, rowCount - bodyEnd - 1) },
        (_, index) => bodyEnd + index + 1,
      ),
    ];
  }, [activeData, activeFilterView, props.epoch.snapshot, props.locale?.id, selection?.active.row]);
  const accessibilityVisualRow = useMemo(
    () => new Map(accessibilityRows.map((logicalRow, visualRow) => [logicalRow, visualRow])),
    [accessibilityRows],
  );
  const accessibilityPresentations = useMemo(
    () =>
      activeDocumentSheet === undefined
        ? null
        : createPresentationResolver({
            document: props.epoch.snapshot.document,
            formulaValues: new Map(
              props.epoch.snapshot.calculation.values.map(({ address, value }) => [address, value]),
            ),
            formulaSpillAnchors: new Map(
              props.epoch.snapshot.calculation.spillAnchors.map(({ address, anchor }) => [
                address,
                anchor,
              ]),
            ),
            cache: accessibilityCache,
            validation: createPresentationValidationResolver(props.epoch.snapshot),
            revisions: {
              document: props.epoch.snapshot.revision,
              calculation: props.epoch.snapshot.calculation.revision,
              condition: props.epoch.snapshot.revision,
              style: props.epoch.snapshot.revision,
              environment: props.epoch.snapshot.revision,
              view: filterViewRevision,
            },
            environment: {
              locale:
                props.locale?.id ??
                props.epoch.snapshot.document.workbook.settings.localeHint ??
                'en-US',
              timeZone: 'UTC',
              dateSystem: props.epoch.snapshot.document.workbook.settings.dateSystem,
              target: 'accessibility',
            },
            ...(activeSheet === null
              ? {}
              : { activeFilterView: props.epoch.controller.getActiveFilterView(activeSheet) }),
          }),
    [
      accessibilityCache,
      activeDocumentSheet,
      activeSheet,
      filterViewRevision,
      props.epoch.controller,
      props.epoch.snapshot,
      props.locale?.id,
    ],
  );
  const accessibilityViewport = (() => {
    void accessibilityViewportRevision;
    const viewport = engineSlot.get()?.interactionSnapshot()?.viewport;
    const visible = viewport === undefined ? null : visibleCellRange(viewport);
    if (visible !== null) return visible;
    if (selection === null) return null;
    return {
      start: {
        row: Math.max(0, selection.active.row - 9),
        column: Math.max(0, selection.active.column - 4),
      },
      end: {
        row: selection.active.row + 10,
        column: selection.active.column + 5,
      },
    };
  })();
  const accessibilityVisualViewport =
    accessibilityViewport === null
      ? null
      : (() => {
          const fallback = accessibilityVisualRow.get(selection?.active.row ?? 0) ?? 0;
          const start = accessibilityVisualRow.get(accessibilityViewport.start.row) ?? fallback;
          const end = accessibilityVisualRow.get(accessibilityViewport.end.row) ?? fallback;
          return { start: Math.min(start, end), end: Math.max(start, end) };
        })();
  const activeStyle =
    selection === null || activeData === null
      ? (initialOptions.defaultStyle ?? {})
      : selectCellStyle(
          activeData,
          selection.active.row,
          selection.active.column,
          initialOptions.defaultStyle,
        );
  const toolbarProps = Object.freeze<ToolbarRenderProps>({
    selection: selection === null ? null : clonePublic(selection),
    activeStyle: clonePublic(activeStyle),
    readOnly: props.readOnly ?? false,
    canUndo: props.epoch.snapshot.canUndo,
    canRedo: props.epoch.snapshot.canRedo,
    merged: mergedSelection(renderRuntime),
    frozen: frozenSheet(renderRuntime),
    disabledActions: readonlySet(disabledToolbarActions(renderRuntime)),
    execute,
  });
  const sheetTabsProps = Object.freeze<SheetTabsRenderProps>({
    sheets: Object.freeze(
      props.epoch.snapshot.sheets.map((sheet) =>
        Object.freeze({
          id: sheet.id,
          index: sheet.index,
          name: sheet.name,
        }),
      ),
    ),
    activeSheet,
    readOnly: props.readOnly ?? false,
    ...tabActions,
  });
  const filterSelection = filterAuthority?.selection ?? selection;
  const filterValues = filterAuthority?.values ?? [];
  const t = createTranslator(props.locale);
  const accessibilityIdPrefix = useId();
  const templateSelection =
    selection === null
      ? undefined
      : {
          sheetId: selection.sheet as unknown as import('../document').DocumentSheetId,
          start: selection.range.start,
          end: selection.range.end,
        };
  const locateTemplateBinding = useCallback(
    (bindingId: import('../document').BindingId) => {
      const binding = props.template?.bindings.find(({ id }) => id === bindingId);
      if (binding === undefined) return;
      const range =
        binding.type === 'value'
          ? {
              sheetId: binding.target.sheetId,
              start: { row: binding.target.row, column: binding.target.column },
              end: { row: binding.target.row, column: binding.target.column },
            }
          : binding.range;
      const sheet = range.sheetId as unknown as SheetId;
      const state = createSelectionState(range.start, range.end);
      const target: Selection = { sheet, range: state.range, active: state.active };
      setActiveSheet(sheet);
      engineSlot.get()?.render(controller.getSnapshot(), sheet);
      engineSlot.get()?.stageSelection(state);
      setSelection(target);
      dispatcher.emitSelectionChange(target);
      engineSlot.get()?.render(controller.getSnapshot(), sheet);
    },
    [controller, dispatcher, engineSlot, props.template, setActiveSheet],
  );
  const templateDecorations = useMemo(() => {
    if (
      props.mode !== 'template' ||
      props.template === undefined ||
      activeDocumentSheet === undefined
    )
      return [];
    const invalidBindings = new Set(
      compileSpreadsheetTemplate(props.epoch.snapshot.document, props.template).diagnostics.flatMap(
        (diagnostic) =>
          diagnostic.severity === 'error' && diagnostic.location?.bindingId !== undefined
            ? [diagnostic.location.bindingId]
            : [],
      ),
    );
    const decorations: import('../engine').TemplateCanvasDecoration[] =
      props.template.bindings.flatMap((binding) => {
        const range =
          binding.type === 'value'
            ? {
                sheetId: binding.target.sheetId,
                start: { row: binding.target.row, column: binding.target.column },
                end: { row: binding.target.row, column: binding.target.column },
              }
            : binding.range;
        return range.sheetId === activeDocumentSheet.id
          ? [
              {
                range: { start: range.start, end: range.end },
                kind: binding.type === 'repeat-rows' ? ('repeat' as const) : ('value' as const),
                label: binding.id,
                invalid: invalidBindings.has(binding.id),
              },
            ]
          : [];
      });
    const profile =
      props.template.printProfiles.find(({ id }) => id === props.activePrintProfileId) ??
      props.template.printProfiles[0];
    if (profile !== undefined) {
      for (const target of profile.targets) {
        const ranges =
          target.type === 'range'
            ? [target.range]
            : target.type === 'ranges'
              ? target.ranges
              : target.sheetId === activeDocumentSheet.id
                ? [
                    {
                      sheetId: activeDocumentSheet.id,
                      start: { row: 0, column: 0 },
                      end: {
                        row: Math.max(0, ...activeDocumentSheet.cells.map(({ row }) => row)),
                        column: Math.max(
                          0,
                          ...activeDocumentSheet.cells.map(({ column }) => column),
                        ),
                      },
                    },
                  ]
                : [];
        for (const range of ranges) {
          if (range.sheetId !== activeDocumentSheet.id) continue;
          decorations.push({
            range: { start: range.start, end: range.end },
            kind: 'print',
            label: profile.name,
          });
        }
      }
    }
    return decorations;
  }, [
    activeDocumentSheet,
    props.activePrintProfileId,
    props.epoch.snapshot.document,
    props.mode,
    props.template,
  ]);
  useEffect(() => {
    const engine = engineSlot.get();
    if (engine === null) return;
    engine.updateTemplateDecorations(templateDecorations);
  }, [engineGeneration, engineSlot, templateDecorations]);
  return (
    <div
      ref={rootCallback}
      className={classNames(props.className)}
      style={props.style}
      data-tego-sheet=""
      data-mode={props.epoch.mode}
      data-grid-visible={props.options?.showGrid === false ? 'false' : 'true'}
      data-context-menu-enabled={props.options?.showContextMenu === false ? 'false' : 'true'}
      role="grid"
      aria-rowcount={activeData === null ? undefined : accessibilityRows.length}
      aria-colcount={
        activeData === null
          ? undefined
          : Math.max(
              selection?.active.column ?? 0,
              typeof activeData.cols?.len === 'number' ? activeData.cols.len - 1 : 0,
            ) + 1
      }
      aria-activedescendant={
        selection === null
          ? undefined
          : `${accessibilityIdPrefix}-r${selection.active.row}-c${selection.active.column}`
      }
      tabIndex={0}
    >
      <CommitAuthority commit={commitRuntime} />
      <SheetChrome
        toolbar={toolbarProps}
        toolbarRenderer={props.toolbar}
        tabs={sheetTabsProps}
        tabsRenderer={props.sheetTabs}
        locale={props.locale}
        editor={editor}
        contextMenu={contextMenu}
        filterColumn={filterSelection?.active.column ?? null}
        filterValues={filterValues}
        filterOpen={filterOpen}
        notification={notification}
        paintFormatActive={paintSource !== null}
        validationOpen={validationOpen}
        onCloseContextMenu={closeContextMenu}
        onCloseFilter={() => {
          setFilterAuthority(null);
          closeFilter();
        }}
        onCloseValidation={closeValidation}
        onDismissNotification={() => setNotification(null)}
        onExecute={execute}
        onExecuteContext={executeContext}
        onFilter={(filter: FilterDefinition) => {
          const authority = filterAuthority;
          setFilterAuthority(null);
          closeFilter();
          const runtime = runtimeAuthority.committed(renderToken);
          if (runtime !== null && authority !== null) {
            executeAction(
              runtime,
              { type: 'set-filter', filter },
              authority.source,
              authority.selection,
            );
          }
        }}
        onOpenFilter={openToolbarFilter}
        onOpenContextFilter={openContextFilter}
        onOpenValidation={openToolbarValidation}
        onOpenContextValidation={openContextValidation}
        onRemoveValidation={() => {
          const authority = validationAuthorityRef.current;
          closeValidation();
          const runtime = runtimeAuthority.committed(renderToken);
          if (runtime !== null && authority !== null) {
            executeAction(
              runtime,
              { type: 'remove-validation' },
              authority.source,
              authority.selection,
            );
          }
        }}
        onValidation={(rule: ValidationRule) => {
          const authority = validationAuthorityRef.current;
          closeValidation();
          const runtime = runtimeAuthority.committed(renderToken);
          if (runtime !== null && authority !== null) {
            executeAction(
              runtime,
              { type: 'set-validation', rule },
              authority.source,
              authority.selection,
            );
          }
        }}
      >
        {sheets.length === 0 ? (
          <EmptyWorkbook readOnly={renderRuntime.readOnly} onAddSheet={addFirstSheet} t={t} />
        ) : (
          <>
            <canvas
              ref={canvasRef}
              className="tego-sheet__canvas"
              data-template-binding-count={
                props.mode === 'template' ? props.template?.bindings.length : undefined
              }
            />
            {selection === null ||
            activeData === null ||
            activeDocumentSheet === undefined ||
            accessibilityViewport === null ||
            accessibilityVisualViewport === null ||
            accessibilityPresentations === null ? null : (
              <div className="tego-sheet__accessibility-grid">
                <AccessibilityGrid
                  rowCount={accessibilityRows.length}
                  rowOrder={accessibilityRows}
                  columnCount={Math.max(
                    selection.active.column + 1,
                    typeof activeData.cols?.len === 'number' ? activeData.cols.len : 0,
                  )}
                  viewport={{
                    rowStart: accessibilityVisualViewport.start,
                    rowEnd: accessibilityVisualViewport.end,
                    columnStart: accessibilityViewport.start.column,
                    columnEnd: accessibilityViewport.end.column,
                  }}
                  activeCell={selection.active}
                  selection={selection.range}
                  editorOpen={editor !== null}
                  embedded
                  idPrefix={accessibilityIdPrefix}
                  readOnly={renderRuntime.readOnly || props.epoch.snapshot.readOnly}
                  restoreFocus={false}
                  resolvePresentation={(point) =>
                    accessibilityPresentations.resolve({
                      sheetId: activeDocumentSheet.id,
                      ...point,
                    })
                  }
                  onActivate={(point) => {
                    if (activeSheet === null) return;
                    const state = createSelectionState(point);
                    const target = {
                      sheet: activeSheet,
                      range: state.range,
                      active: state.active,
                    };
                    engineSlot.get()?.stageSelection(state);
                    setSelection(target);
                    dispatcher.emitSelectionChange(target);
                    engineSlot.get()?.render(controller.getSnapshot(), activeSheet);
                  }}
                  onRequestEdit={(point) => requestEdit(point, undefined, 'pointer')}
                />
              </div>
            )}
            {visibleObjectProjections.length === 0 ? null : (
              <div className="tego-sheet__accessibility-grid">
                <AccessibilityObjects
                  objects={visibleObjectProjections}
                  selectedObjectId={activeSelectedObjectId}
                  readOnly={renderRuntime.readOnly || props.epoch.snapshot.readOnly}
                  onSelect={(objectId) => {
                    if (activeSheet !== null) setSelectedObject({ sheet: activeSheet, objectId });
                    engineSlot.get()?.setSelectedObject(objectId);
                  }}
                  onChange={(object) => {
                    if (
                      activeSheet === null ||
                      renderRuntime.readOnly ||
                      props.epoch.snapshot.readOnly
                    )
                      return;
                    dispatcher.dispatchUi(
                      { type: 'set-sheet-object', sheet: activeSheet, object },
                      'keyboard',
                    );
                  }}
                />
              </div>
            )}
          </>
        )}
      </SheetChrome>
      {props.mode === 'template' && props.template !== undefined ? (
        <TemplateDesignerSurface
          document={props.epoch.snapshot.document}
          template={props.template}
          onChange={props.onTemplateChange ?? (() => undefined)}
          onDiagnostics={props.onDiagnostics}
          selection={templateSelection}
          onLocateBinding={locateTemplateBinding}
          activeProfileId={props.activePrintProfileId}
          onActiveProfileChange={props.onActivePrintProfileChange}
        />
      ) : null}
    </div>
  );
}

const ForwardedRuntime = forwardRef(Runtime);

function PreviewRenderSession(props: {
  readonly compiled: NonNullable<ReturnType<typeof compileSpreadsheetTemplate>['template']>;
  readonly profile: NonNullable<TegoSheetProps['template']>['printProfiles'][number];
  readonly sampleData: unknown;
  readonly environment: NonNullable<TegoSheetProps['renderEnvironment']>;
  readonly onDiagnostics: TegoSheetProps['onDiagnostics'];
}) {
  const [result, setResult] = useState<Awaited<
    ReturnType<typeof renderSpreadsheetTemplate>
  > | null>(null);
  const { compiled, environment, onDiagnostics, profile, sampleData } = props;
  useEffect(() => {
    const controller = new AbortController();
    void renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: sampleData,
        profileId: profile.id,
        missingValue: 'warning-and-blank',
        signal: controller.signal,
      },
      environment,
    ).then((next) => {
      if (controller.signal.aborted) return;
      onDiagnostics?.(next.diagnostics);
      setResult(next);
    });
    return () => controller.abort();
  }, [compiled, environment, onDiagnostics, profile.id, sampleData]);
  if (result?.document !== undefined) {
    const scale =
      profile.page.scale.type === 'fixed'
        ? `Fixed ${profile.page.scale.value}`
        : profile.page.scale.type === 'fit-width'
          ? `Fit width ${profile.page.scale.pages}`
          : 'Fit page';
    return (
      <section aria-label="Template preview">
        <p aria-label="Template preview metadata">
          {result.document.print.pages.length} page(s) · {profile.page.paper.type} ·{' '}
          {profile.page.orientation} · {scale}
        </p>
        <TemplatePreview document={result.document} />
        <section aria-label="Template preview diagnostics" aria-live="polite">
          {result.diagnostics.map((diagnostic, index) => (
            <p key={`${diagnostic.code}-${index}`}>{diagnostic.message}</p>
          ))}
        </section>
      </section>
    );
  }
  return (
    <section aria-label="Template preview diagnostics" aria-live="polite">
      {result?.diagnostics.map((diagnostic, index) => (
        <p key={`${diagnostic.code}-${index}`}>{diagnostic.message}</p>
      ))}
    </section>
  );
}

function TemplateSurface(props: {
  readonly document: import('../document').SpreadsheetDocument;
  readonly template: NonNullable<TegoSheetProps['template']>;
  readonly sampleData: unknown;
  readonly environment: NonNullable<TegoSheetProps['renderEnvironment']>;
  readonly onDiagnostics: TegoSheetProps['onDiagnostics'];
  readonly activeProfileId?: string;
}) {
  const { activeProfileId, document, environment, onDiagnostics, sampleData, template } = props;
  const compilation = useMemo(
    () => compileSpreadsheetTemplate(document, template),
    [document, template],
  );
  useEffect(() => {
    onDiagnostics?.(compilation.diagnostics);
  }, [compilation.diagnostics, onDiagnostics]);
  const profile =
    template.printProfiles.find(({ id }) => id === activeProfileId) ?? template.printProfiles[0];
  if (compilation.template !== undefined && profile !== undefined) {
    return (
      <PreviewRenderSession
        key={`${compilation.template.sourceDocumentHash}:${template.id}:${profile.id}`}
        compiled={compilation.template}
        profile={profile}
        sampleData={sampleData}
        environment={environment}
        onDiagnostics={onDiagnostics}
      />
    );
  }
  return (
    <section aria-label="Template preview diagnostics" aria-live="polite">
      {compilation.diagnostics.map((diagnostic, index) => (
        <p key={`${diagnostic.code}-${index}`}>{diagnostic.message}</p>
      ))}
    </section>
  );
}

function TemplateDesignerSurface(props: {
  readonly document: import('../document').SpreadsheetDocument;
  readonly template: NonNullable<TegoSheetProps['template']>;
  readonly onChange: NonNullable<TegoSheetProps['onTemplateChange']>;
  readonly onDiagnostics: TegoSheetProps['onDiagnostics'];
  readonly selection?: import('../document').DocumentCellRange;
  readonly onLocateBinding?: (bindingId: import('../document').BindingId) => void;
  readonly activeProfileId?: string;
  readonly onActiveProfileChange?: (profileId: string) => void;
}) {
  const {
    activeProfileId,
    document,
    onActiveProfileChange,
    onChange,
    onDiagnostics,
    onLocateBinding,
    selection,
    template,
  } = props;
  const diagnostics = useMemo(
    () => compileSpreadsheetTemplate(document, template).diagnostics,
    [document, template],
  );
  useEffect(() => onDiagnostics?.(diagnostics), [diagnostics, onDiagnostics]);
  return (
    <TemplateDesigner
      template={template}
      diagnostics={diagnostics}
      onChange={onChange}
      selection={selection}
      onLocateBinding={onLocateBinding}
      activeProfileId={activeProfileId}
      onActiveProfileChange={onActiveProfileChange}
    />
  );
}

/**
 * Renders an interactive spreadsheet with controlled or uncontrolled document ownership.
 * Use a {@link TegoSheetHandle} ref for cell, sheet, validation, and layout operations.
 *
 * @example
 * ```tsx
 * import { useState } from 'react';
 * import { createSpreadsheetDocument, TegoSheet, type SpreadsheetDocument } from 'tego-sheet';
 *
 * function BudgetSheet() {
 *   const [document, setDocument] = useState<SpreadsheetDocument>(
 *     createSpreadsheetDocument({ sheetName: 'Budget' }),
 *   );
 *   return <TegoSheet document={document} onDocumentChange={setDocument} />;
 * }
 * ```
 */
export const TegoSheet = forwardRef<TegoSheetHandle, TegoSheetProps>(
  function TegoSheet(props, ref) {
    const mountOptions = useMountOptionWarnings(props.initialActiveSheetIndex, props.options);
    const [mountActiveSheetIndex] = useState(() => props.initialActiveSheetIndex);
    const epoch = useControllerEpoch(props);
    const controlled = useControlledWorkbook({
      epoch,
      controlledDocument: props.document,
      onError: props.onError,
    });
    if (epoch === null) {
      return (
        <div
          className={classNames(props.className)}
          style={props.style}
          data-tego-sheet=""
          data-mode="initializing"
        />
      );
    }
    if (props.mode === 'preview') {
      if (props.template === undefined || props.renderEnvironment === undefined) {
        return (
          <section aria-label="Template preview diagnostics">
            Preview mode requires a template and deterministic render environment.
          </section>
        );
      }
      return (
        <TemplateSurface
          document={epoch.snapshot.document}
          template={props.template}
          sampleData={props.sampleData}
          environment={props.renderEnvironment}
          onDiagnostics={props.onDiagnostics}
          activeProfileId={props.activePrintProfileId}
        />
      );
    }
    return (
      <ForwardedRuntime
        {...props}
        controlled={controlled}
        epoch={epoch}
        mountOptions={mountOptions}
        mountActiveSheetIndex={mountActiveSheetIndex}
        ref={ref}
      />
    );
  },
);

TegoSheet.displayName = 'TegoSheet';
