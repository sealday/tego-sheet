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
    expect(edit).toContain('input.controller.transact(');
    expect(edit).not.toContain('input.controller.execute(');
  });
});
