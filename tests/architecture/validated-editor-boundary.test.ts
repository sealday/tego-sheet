import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const root = resolve(import.meta.dirname, '../..');

describe('validated editor mutation boundary', () => {
  it('routes typed validation through the async transaction path before sync dispatch', () => {
    const hook = readFileSync(resolve(root, 'src/react/hooks/use-cell-editor-runtime.ts'), 'utf8');
    const validationBranch = hook.indexOf('documentValidationRequest(');
    const asyncDispatch = hook.indexOf('.dispatchValidatedUi(', validationBranch);
    const syncDispatch = hook.indexOf('.dispatcher.dispatchUi(', validationBranch);
    expect(validationBranch).toBeGreaterThan(-1);
    expect(asyncDispatch).toBeGreaterThan(validationBranch);
    expect(syncDispatch).toBeGreaterThan(asyncDispatch);
    expect(hook).toContain('prevalidated/internal synchronous dispatch path');
  });

  it('keeps accepted validation mutations on one revision-bound transaction', () => {
    const edit = readFileSync(resolve(root, 'src/validation/edit.ts'), 'utf8');
    expect(edit).toContain('const baseRevision = input.controller.getSnapshot().revision');
    expect(edit).toContain('executeValidatedTransaction(');
    expect(edit).not.toContain('commitValidatedDocumentTransaction(');
    expect(edit).not.toContain('input.controller.transact(');
    expect(edit).not.toContain('input.controller.execute(');
  });

  it('routes the imperative cell edit entry through the validated transaction path', () => {
    const handle = readFileSync(resolve(root, 'src/react/hooks/use-tego-sheet-handle.ts'), 'utf8');
    const entry = handle.indexOf('setCellText(address, text)');
    const request = handle.indexOf('documentValidationRequest(', entry);
    const validatedDispatch = handle.indexOf('.dispatchValidatedUi(', request);
    const directDispatch = handle.indexOf('.dispatcher.dispatchRef(', entry);
    expect(entry).toBeGreaterThan(-1);
    expect(request).toBeGreaterThan(entry);
    expect(validatedDispatch).toBeGreaterThan(request);
    expect(directDispatch).toBeGreaterThan(entry);
    expect(handle.slice(entry, validatedDispatch)).toContain(
      'if (unresolvedRequest === undefined)',
    );
  });
});
