import { useImperativeHandle, useLayoutEffect, useState, type ForwardedRef } from 'react';
import {
  selectCell,
  selectCellStyle,
  TegoSheetException,
  type CellAddress,
  type CellStyle,
  type ChangeSource,
  type SheetId,
} from '../../core';
import type { WorkbookCommand } from '../../core/commands/workbook-command';
import type { ControllerEpoch } from './use-controller-epoch';
import type { EventDispatcher } from '../adapters/event-dispatcher';
import type { EngineAdapterSlot } from './use-canvas-engine';
import type { TegoSheetHandle } from '../tego-sheet.types';
import type {
  ValidationEngineOptions,
  ValidationResult as AdvancedValidationResult,
} from '../../validation';
import { documentValidationRequest } from '../../validation/document-rule';
import { beginCellValidation } from '../../validation/edit-gate';

function invalid(message: string): TegoSheetException {
  return new TegoSheetException({
    code: 'INVALID_COMMAND',
    message,
    recoverable: true,
  });
}

function committedResult(
  dispatcher: EventDispatcher,
  command: WorkbookCommand,
  source: ChangeSource,
): unknown {
  const outcome = dispatcher.dispatchRef(command, source);
  return outcome.status === 'committed' ? outcome.commit.result : undefined;
}

export interface TegoSheetHandleRuntime {
  readonly activeSheet: SheetId | null;
  readonly controller: ControllerEpoch['controller'];
  readonly defaultStyle: CellStyle | undefined;
  readonly dispatcher: EventDispatcher;
  readonly engineSlot: EngineAdapterSlot;
  readonly isActive: () => boolean;
  readonly root: HTMLDivElement | null;
  readonly setActiveSheet: (sheet: SheetId | null) => void;
  readonly refreshFilterView?: () => void;
  readonly validation: ValidationEngineOptions;
  readonly confirmValidationWarning?: (
    result: AdvancedValidationResult,
  ) => boolean | Promise<boolean>;
}

export interface RuntimeCapture<Runtime extends TegoSheetHandleRuntime> {
  readonly activeDecisionVersion: number;
  readonly runtime: Runtime;
}

export interface TegoSheetRuntimeAuthority<Runtime extends TegoSheetHandleRuntime> {
  readonly capture: () => RuntimeCapture<Runtime>;
  readonly commit: (token: object, runtime: Runtime) => void;
  readonly committed: (token: object) => Runtime | null;
  readonly compareAndSetActiveSheet: (
    capture: RuntimeCapture<Runtime>,
    sheet: SheetId | null,
  ) => boolean;
  readonly deactivate: () => void;
  readonly patchRoot: (root: HTMLDivElement) => boolean;
  readonly require: () => Runtime;
  readonly activate: (sheet: SheetId | null) => void;
  readonly trackValidation: (controller: { readonly abort: () => void }) => () => void;
}

function createRuntimeAuthority<
  Runtime extends TegoSheetHandleRuntime,
>(): TegoSheetRuntimeAuthority<Runtime> {
  let current: Runtime | null = null;
  let activeDecisionVersion = 0;
  const committedTokens = new WeakSet<object>();
  const pendingValidations = new Set<{ readonly abort: () => void }>();
  const requireRuntime = () => {
    if (current === null || !current.isActive()) {
      throw invalid('TegoSheet handle runtime is inactive');
    }
    return current;
  };
  const applyActiveSheet = (runtime: Runtime, sheet: SheetId | null) => {
    activeDecisionVersion += 1;
    current = { ...runtime, activeSheet: sheet };
    runtime.setActiveSheet(sheet);
  };
  return {
    capture: () => ({ activeDecisionVersion, runtime: requireRuntime() }),
    commit(token, runtime) {
      if (
        current !== null &&
        (current.controller !== runtime.controller || current.activeSheet !== runtime.activeSheet)
      )
        activeDecisionVersion += 1;
      current = runtime;
      committedTokens.add(token);
    },
    committed(token) {
      if (!committedTokens.has(token)) return null;
      const runtime = current;
      return runtime?.isActive() === true ? runtime : null;
    },
    compareAndSetActiveSheet(capture, sheet) {
      const runtime = current;
      if (
        runtime === null ||
        !runtime.isActive() ||
        runtime.controller !== capture.runtime.controller ||
        activeDecisionVersion !== capture.activeDecisionVersion
      )
        return false;
      applyActiveSheet(runtime, sheet);
      return true;
    },
    deactivate() {
      activeDecisionVersion += 1;
      current = null;
      for (const controller of pendingValidations) controller.abort();
      pendingValidations.clear();
    },
    patchRoot(root) {
      const runtime = current;
      if (runtime === null || !runtime.isActive()) return false;
      current = { ...runtime, root };
      return true;
    },
    require: requireRuntime,
    activate(sheet) {
      applyActiveSheet(requireRuntime(), sheet);
    },
    trackValidation(controller) {
      pendingValidations.add(controller);
      return () => pendingValidations.delete(controller);
    },
  };
}

function runtimeSheet(runtime: TegoSheetHandleRuntime, address: CellAddress) {
  const snapshot = runtime.controller.getSnapshot();
  const index = snapshot.sheets.findIndex((sheet) => sheet.id === address.sheet);
  if (index < 0) throw invalid(`Unknown sheet ID: ${address.sheet}`);
  return snapshot.projection[index]!;
}

function createStableHandle<Runtime extends TegoSheetHandleRuntime>(
  authority: TegoSheetRuntimeAuthority<Runtime>,
): TegoSheetHandle {
  return {
    focus() {
      authority.require().root?.focus();
    },
    getDocument: () => authority.require().controller.getDocument(),
    getCell(address) {
      const runtime = authority.require();
      runtime.controller.getCellText(address);
      return selectCell(runtimeSheet(runtime, address), address.row, address.column);
    },
    getCellStyle(address) {
      const runtime = authority.require();
      runtime.controller.getCellText(address);
      return selectCellStyle(
        runtimeSheet(runtime, address),
        address.row,
        address.column,
        runtime.defaultStyle,
      );
    },
    setCellText(address, text) {
      const runtime = authority.require();
      const validationAddress = {
        sheetId: address.sheet as unknown as import('../../document').DocumentSheetId,
        row: address.row,
        column: address.column,
      };
      const unresolvedRequest = documentValidationRequest(
        runtime.controller.getDocument(),
        validationAddress,
        text,
      );
      if (unresolvedRequest === undefined) {
        runtime.dispatcher.dispatchRef({ type: 'set-cell-text', address, text }, 'ref');
        return;
      }
      const lease = beginCellValidation(runtime.controller, validationAddress);
      const untrack = authority.trackValidation(lease);
      void runtime.dispatcher
        .dispatchValidatedUi(
          {
            address: validationAddress,
            text,
            validation: runtime.validation,
            signal: lease.signal,
            ...(runtime.confirmValidationWarning === undefined
              ? {}
              : { confirmWarning: runtime.confirmValidationWarning }),
            canCommit: () =>
              runtime.isActive() &&
              lease.isCurrent() &&
              runtime.controller === authority.require().controller,
          },
          'ref',
        )
        .then((outcome) => {
          if (outcome.status === 'rejected' && runtime.isActive()) {
            runtime.dispatcher.reportUiError(outcome.error);
          }
        })
        .finally(() => {
          untrack();
          lease.release();
        });
    },
    addSheet(name) {
      const capture = authority.capture();
      const { runtime } = capture;
      const wasEmpty = runtime.controller.getSnapshot().sheets.length === 0;
      const result = committedResult(
        runtime.dispatcher,
        name === undefined ? { type: 'add-sheet' } : { type: 'add-sheet', name },
        'ref',
      );
      if (typeof result !== 'string') throw invalid('Adding a sheet did not return a sheet ID');
      const sheet = result as SheetId;
      if (wasEmpty) authority.compareAndSetActiveSheet(capture, sheet);
      return sheet;
    },
    deleteSheet(sheet) {
      const capture = authority.capture();
      const { runtime } = capture;
      const before = runtime.controller.getSnapshot();
      const removedIndex = before.sheets.findIndex((item) => item.id === sheet);
      runtime.dispatcher.dispatchRef({ type: 'delete-sheet', sheet }, 'ref');
      if (runtime.activeSheet !== sheet) return;
      const after = runtime.controller.getSnapshot();
      const replacementIndex = Math.min(removedIndex, after.sheets.length - 1);
      authority.compareAndSetActiveSheet(
        capture,
        replacementIndex < 0 ? null : after.sheets[replacementIndex]!.id,
      );
    },
    renameSheet(sheet, name) {
      authority.require().dispatcher.dispatchRef({ type: 'rename-sheet', sheet, name }, 'ref');
    },
    activateSheet(sheet) {
      const runtime = authority.require();
      const index = runtime.controller.getSnapshot().sheets.findIndex((item) => item.id === sheet);
      if (index < 0) throw invalid(`Unknown sheet ID: ${sheet}`);
      authority.activate(sheet);
      runtime.dispatcher.emitActiveSheetChange({ sheet, index, source: 'ref' });
    },
    activateFilterView(sheet, viewId) {
      const runtime = authority.require();
      runtime.controller.activateFilterView(sheet, viewId);
      runtime.refreshFilterView?.();
    },
    deactivateFilterView(sheet) {
      const runtime = authority.require();
      runtime.controller.deactivateFilterView(sheet);
      runtime.refreshFilterView?.();
    },
    undo() {
      authority.require().dispatcher.dispatchRef({ type: 'undo' }, 'ref');
    },
    redo() {
      authority.require().dispatcher.dispatchRef({ type: 'redo' }, 'ref');
    },
    validate: () => authority.require().controller.validate(),
    recalculateLayout() {
      authority.require().engineSlot.get()?.recalculateLayout();
    },
  };
}

export function useTegoSheetHandle<Runtime extends TegoSheetHandleRuntime>(
  forwardedRef: ForwardedRef<TegoSheetHandle>,
): TegoSheetRuntimeAuthority<Runtime> {
  const [authority] = useState(createRuntimeAuthority<Runtime>);
  const [handle] = useState(() => createStableHandle(authority));
  useLayoutEffect(() => () => authority.deactivate(), [authority]);
  useImperativeHandle(forwardedRef, () => handle, [handle]);
  return authority;
}
