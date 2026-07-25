# Output Studio Playground Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use
> superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement
> this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a complete Output Studio workspace to the documentation Playground that demonstrates
template editing, deterministic generation, exact preview, browser printing, and PDF/PNG/XLSX
output.

**Architecture:** Preserve the existing spreadsheet presets behind a new workspace shell. Build
Output Studio from focused fixture, model, pipeline, preview, workbench, and output-action modules;
all output actions consume one committed `GeneratedDocument`, while draft edits remain explicitly
stale until regeneration.

**Tech Stack:** React 19, TypeScript 6, Docusaurus 3, Vitest, Testing Library, Playwright, existing
tego-sheet template and output adapters.

## Global Constraints

- No new dependencies.
- No changes to public printing or output APIs.
- Existing Spreadsheet preset behavior remains stable.
- Use public tego-sheet exports wherever a public surface exists.
- Automated browser tests must not open a real system print dialog.
- Initial Output Studio content is a working deterministic invoice example.
- Draft changes disable every output until `Apply & regenerate` succeeds.
- Preview, print, PDF, and PNG consume the same generated page display list.
- XLSX consumes the semantic workbook in the same `GeneratedDocument`.
- All interactive controls remain at least 44 CSS pixels tall.

---

## File structure

### New files

- `website/src/components/playground/playground-workspace.ts`
  - Parses and serializes top-level workspace URL state.
- `website/src/components/playground/output-studio-model.ts`
  - Pure revision, dirty, rendering, blocked, export, and filename logic.
- `website/src/components/playground/output-studio-fixtures.ts`
  - Prepared invoice document, template, sample data, and render environment.
- `website/src/components/playground/output-studio-pipeline.ts`
  - Compiles and renders one abortable committed revision.
- `website/src/components/playground/output-download.ts`
  - Browser Blob download lifecycle and deterministic cleanup.
- `website/src/components/playground/output-studio.tsx`
  - Output Studio composition, template workbench, preview, diagnostics, and actions.
- `tests/unit/website/playground-workspace.test.ts`
  - Workspace URL contract.
- `tests/unit/website/output-studio-model.test.ts`
  - Pure pipeline state and filenames.
- `tests/unit/website/output-studio-pipeline.test.ts`
  - Compile/render and stale request behavior.
- `tests/component/docs-output-studio.test.tsx`
  - Output Studio interactions and adapter contracts.

### Modified files

- `website/src/components/playground/playground.tsx`
  - Becomes the workspace shell and delegates existing content to `SpreadsheetWorkspace`.
- `website/src/components/playground/playground.module.css`
  - Adds workspace navigation and responsive Output Studio layout.
- `tests/component/docs-playground.test.tsx`
  - Locks existing Spreadsheet behavior behind the new shell.
- `tests/docs/docs.spec.ts`
  - Adds direct URL, regeneration, diagnostic, and keyboard coverage.
- `tests/docs-visual/docs-visual.spec.ts`
  - Adds wide, intermediate, narrow, stale, and blocked screenshots.
- `website/docs/guides/printing.md`
  - Links the runnable Output Studio demonstration.

---

### Task 1: Workspace URL contract

**Files:**

- Create: `website/src/components/playground/playground-workspace.ts`
- Create: `tests/unit/website/playground-workspace.test.ts`
- Modify: `website/src/components/playground/playground-model.ts`

**Interfaces:**

- Produces:
  - `type PlaygroundWorkspace = 'spreadsheet' | 'output'`
  - `readPlaygroundLocation(search: string): PlaygroundLocation`
  - `writePlaygroundLocation(pathname: string, search: string, next: PlaygroundLocation): string`
- Consumes: existing `PlaygroundMode` and `parsePlaygroundMode`.

- [ ] **Step 1: Write failing workspace parsing tests**

```ts
import { describe, expect, it } from 'vitest';
import {
  readPlaygroundLocation,
  writePlaygroundLocation,
} from '../../../website/src/components/playground/playground-workspace';

describe('Playground workspace URL state', () => {
  it('keeps legacy mode links in the Spreadsheet workspace', () => {
    expect(readPlaygroundLocation('?mode=controlled')).toEqual({
      workspace: 'spreadsheet',
      mode: 'controlled',
    });
  });

  it('selects Output Studio without discarding the remembered spreadsheet mode', () => {
    expect(readPlaygroundLocation('?workspace=output&mode=locales')).toEqual({
      workspace: 'output',
      mode: 'locales',
    });
  });

  it('canonicalizes invalid values and preserves unrelated parameters', () => {
    expect(
      writePlaygroundLocation('/tego-sheet/playground', '?theme=dark&workspace=nope&mode=nope', {
        workspace: 'spreadsheet',
        mode: 'uncontrolled',
      }),
    ).toBe('/tego-sheet/playground?theme=dark&workspace=spreadsheet&mode=uncontrolled');
  });
});
```

- [ ] **Step 2: Run the tests and verify RED**

Run:

```bash
npx vitest run --project unit tests/unit/website/playground-workspace.test.ts
```

Expected: FAIL because `playground-workspace.ts` does not exist.

- [ ] **Step 3: Implement the URL model**

```ts
import { parsePlaygroundMode, type PlaygroundMode } from './playground-model';

export const PLAYGROUND_WORKSPACES = Object.freeze(['spreadsheet', 'output'] as const);
export type PlaygroundWorkspace = (typeof PLAYGROUND_WORKSPACES)[number];

export interface PlaygroundLocation {
  readonly workspace: PlaygroundWorkspace;
  readonly mode: PlaygroundMode;
}

export function readPlaygroundLocation(search: string): PlaygroundLocation {
  const params = new URLSearchParams(search);
  const workspace = PLAYGROUND_WORKSPACES.includes(
    params.get('workspace') as PlaygroundWorkspace,
  )
    ? (params.get('workspace') as PlaygroundWorkspace)
    : 'spreadsheet';
  return Object.freeze({ workspace, mode: parsePlaygroundMode(params.get('mode')) });
}

export function writePlaygroundLocation(
  pathname: string,
  search: string,
  next: PlaygroundLocation,
): string {
  const params = new URLSearchParams(search);
  params.set('workspace', next.workspace);
  params.set('mode', next.mode);
  return `${pathname}?${params.toString()}`;
}
```

- [ ] **Step 4: Run workspace and existing model tests**

Run:

```bash
npx vitest run --project unit tests/unit/website/playground-workspace.test.ts tests/unit/website/playground-model.test.ts
```

Expected: both files PASS.

- [ ] **Step 5: Commit**

```bash
git add website/src/components/playground/playground-workspace.ts \
  website/src/components/playground/playground-model.ts \
  tests/unit/website/playground-workspace.test.ts
git commit -m "feat(docs): model playground workspaces"
```

---

### Task 2: Output Studio state model

**Files:**

- Create: `website/src/components/playground/output-studio-model.ts`
- Create: `tests/unit/website/output-studio-model.test.ts`

**Interfaces:**

- Produces:
  - `OutputStudioState`
  - `OutputStudioAction`
  - `reduceOutputStudioState(state, action)`
  - `outputFilename(kind, invoiceId, page?)`
  - `hasBlockingDiagnostics(diagnostics)`
- Consumes: public `Diagnostic` and `GeneratedDocument` types.

- [ ] **Step 1: Write failing reducer tests**

```ts
import { describe, expect, it } from 'vitest';
import {
  createOutputStudioState,
  outputFilename,
  reduceOutputStudioState,
} from '../../../website/src/components/playground/output-studio-model';

describe('Output Studio model', () => {
  it('marks output stale after a draft edit and disables generated output', () => {
    const ready = {
      ...createOutputStudioState(),
      phase: 'ready' as const,
      committedRevision: 1,
      generatedRevision: 1,
    };
    expect(reduceOutputStudioState(ready, { type: 'draft-changed' })).toMatchObject({
      phase: 'dirty',
      committedRevision: 1,
      generatedRevision: 1,
    });
  });

  it('ignores a render result from an older revision', () => {
    const rendering = {
      ...createOutputStudioState(),
      phase: 'rendering' as const,
      committedRevision: 3,
    };
    expect(
      reduceOutputStudioState(rendering, {
        type: 'render-succeeded',
        revision: 2,
        document: {} as never,
        diagnostics: [],
      }),
    ).toBe(rendering);
  });

  it('creates deterministic output filenames', () => {
    expect(outputFilename('pdf', 'INV-2026-042')).toBe('invoice-INV-2026-042.pdf');
    expect(outputFilename('png', 'INV-2026-042', 0)).toBe(
      'invoice-INV-2026-042-page-1.png',
    );
    expect(outputFilename('xlsx', 'INV-2026-042')).toBe('invoice-INV-2026-042.xlsx');
  });
});
```

- [ ] **Step 2: Run the model test and verify RED**

Run:

```bash
npx vitest run --project unit tests/unit/website/output-studio-model.test.ts
```

Expected: FAIL because the model module does not exist.

- [ ] **Step 3: Implement the minimal state machine**

```ts
import type { Diagnostic, GeneratedDocument } from 'tego-sheet';

export type OutputStudioPhase = 'ready' | 'dirty' | 'rendering' | 'blocked';
export type OutputKind = 'print' | 'pdf' | 'png' | 'xlsx';

export interface OutputStudioState {
  readonly phase: OutputStudioPhase;
  readonly committedRevision: number;
  readonly generatedRevision: number | null;
  readonly generatedDocument: GeneratedDocument | null;
  readonly diagnostics: readonly Diagnostic[];
  readonly activeOutput: OutputKind | null;
  readonly outputMessage: string;
}

export type OutputStudioAction =
  | { readonly type: 'draft-changed' }
  | { readonly type: 'render-started'; readonly revision: number }
  | {
      readonly type: 'render-succeeded';
      readonly revision: number;
      readonly document: GeneratedDocument;
      readonly diagnostics: readonly Diagnostic[];
    }
  | {
      readonly type: 'render-blocked';
      readonly revision: number;
      readonly diagnostics: readonly Diagnostic[];
    }
  | { readonly type: 'output-started'; readonly kind: OutputKind }
  | { readonly type: 'output-finished'; readonly message: string }
  | { readonly type: 'output-failed'; readonly message: string };

export function createOutputStudioState(): OutputStudioState {
  return {
    phase: 'rendering',
    committedRevision: 0,
    generatedRevision: null,
    generatedDocument: null,
    diagnostics: [],
    activeOutput: null,
    outputMessage: '',
  };
}

export function reduceOutputStudioState(
  state: OutputStudioState,
  action: OutputStudioAction,
): OutputStudioState {
  if (action.type === 'render-started') {
    return {
      ...state,
      phase: 'rendering',
      committedRevision: action.revision,
      activeOutput: null,
      outputMessage: '',
    };
  }
  if ('revision' in action && action.revision !== state.committedRevision) return state;
  if (action.type === 'draft-changed' && state.phase === 'dirty') return state;
  switch (action.type) {
    case 'draft-changed':
      return { ...state, phase: 'dirty', activeOutput: null };
    case 'render-succeeded':
      return {
        ...state,
        phase: 'ready',
        generatedRevision: action.revision,
        generatedDocument: action.document,
        diagnostics: action.diagnostics,
      };
    case 'render-blocked':
      return { ...state, phase: 'blocked', diagnostics: action.diagnostics };
    case 'output-started':
      return { ...state, activeOutput: action.kind, outputMessage: '' };
    case 'output-finished':
    case 'output-failed':
      return { ...state, activeOutput: null, outputMessage: action.message };
  }
}

export function outputFilename(
  kind: Exclude<OutputKind, 'print'>,
  invoiceId: string,
  page = 0,
): string {
  const safeId = invoiceId.replace(/[^A-Za-z0-9._-]+/g, '-');
  return kind === 'png'
    ? `invoice-${safeId}-page-${page + 1}.png`
    : `invoice-${safeId}.${kind}`;
}

export function hasBlockingDiagnostics(diagnostics: readonly Diagnostic[]): boolean {
  return diagnostics.some(({ severity }) => severity === 'error');
}
```

- [ ] **Step 4: Run model tests and correct reducer ordering**

Run:

```bash
npx vitest run --project unit tests/unit/website/output-studio-model.test.ts
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add website/src/components/playground/output-studio-model.ts \
  tests/unit/website/output-studio-model.test.ts
git commit -m "feat(docs): add output studio state model"
```

---

### Task 3: Deterministic invoice fixture and render pipeline

**Files:**

- Create: `website/src/components/playground/output-studio-fixtures.ts`
- Create: `website/src/components/playground/output-studio-pipeline.ts`
- Create: `tests/unit/website/output-studio-pipeline.test.ts`

**Interfaces:**

- Produces:
  - `createInvoiceOutputFixture(): InvoiceOutputFixture`
  - `renderOutputRevision(request): Promise<OutputRevisionResult>`
- Consumes:
  - `compileSpreadsheetTemplate`
  - `hashSpreadsheetDocument`
  - `renderSpreadsheetTemplate`
  - `createFontMetrics`

- [ ] **Step 1: Write the failing fixture and render test**

```ts
import { describe, expect, it } from 'vitest';
import {
  createInvoiceOutputFixture,
} from '../../../website/src/components/playground/output-studio-fixtures';
import {
  renderOutputRevision,
} from '../../../website/src/components/playground/output-studio-pipeline';

describe('Output Studio rendering', () => {
  it('renders the prepared invoice into two deterministic pages', async () => {
    const fixture = createInvoiceOutputFixture();
    const result = await renderOutputRevision({
      revision: 1,
      document: fixture.document,
      template: fixture.template,
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result.revision).toBe(1);
    expect(result.document?.print.pages).toHaveLength(2);
    expect(result.document?.metadata.generatedAt).toBe('2026-07-25T00:00:00.000Z');
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('returns blocking diagnostics instead of an artifact for invalid bindings', async () => {
    const fixture = createInvoiceOutputFixture();
    const result = await renderOutputRevision({
      revision: 2,
      document: fixture.document,
      template: {
        ...fixture.template,
        bindings: [{ ...fixture.template.bindings[0]!, expression: 'missing.customer' }],
      },
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });
    expect(result.document).toBeUndefined();
    expect(result.diagnostics.some(({ severity }) => severity === 'error')).toBe(true);
  });
});
```

- [ ] **Step 2: Run the pipeline test and verify RED**

Run:

```bash
npx vitest run --project unit tests/unit/website/output-studio-pipeline.test.ts
```

Expected: FAIL because both modules are missing.

- [ ] **Step 3: Implement the prepared fixture**

Use `migrateLegacyWorkbook` to create a source document with stable IDs and content on both sides of
the manual break:

```ts
const migrated = migrateLegacyWorkbook(
  [
    {
      name: 'Invoice',
      rows: {
        len: 30,
        0: { cells: { 0: { text: 'INVOICE' }, 4: { text: 'Invoice number' } } },
        2: { cells: { 0: { text: 'Customer' }, 4: { text: 'Number' } } },
        4: {
          cells: {
            0: { text: 'Description' },
            3: { text: 'Quantity' },
            4: { text: 'Unit price' },
            5: { text: 'Amount' },
          },
        },
        6: { cells: { 0: { text: 'Hosting' }, 3: { text: '1' }, 5: { text: '29' } } },
        7: { cells: { 0: { text: 'Support' }, 3: { text: '4' }, 5: { text: '75' } } },
        8: { cells: { 0: { text: 'Training' }, 3: { text: '2' }, 5: { text: '240' } } },
        14: { cells: { 0: { text: 'Terms and payment details' } } },
        18: { cells: { 0: { text: 'Thank you for your business.' } } },
      },
      cols: { len: 6 },
    },
  ],
  {
    ids: {
      documentId: () => 'invoice-document',
      sheetId: () => 'invoice-sheet',
    },
  },
);
if (!migrated.ok) {
  throw new TypeError(migrated.diagnostics.map(({ message }) => message).join('\n'));
}
const document = migrated.document;
```

Create a template with:

```ts
const template: SpreadsheetTemplate = {
  id: 'invoice-template' as never,
  name: 'Customer invoice',
  bindings: [
    {
      id: 'customer-name' as never,
      type: 'value',
      target: { sheetId: 'invoice-sheet' as never, row: 2, column: 0 },
      expression: 'customer.name',
    },
    {
      id: 'invoice-number' as never,
      type: 'value',
      target: { sheetId: 'invoice-sheet' as never, row: 2, column: 4 },
      expression: 'invoice.id',
    },
  ],
  printProfiles: [
    {
      id: 'invoice-a4',
      name: 'Invoice · A4',
      targets: [{ type: 'sheet', sheetId: 'invoice-sheet' as never }],
      page: {
        paper: { type: 'A4' },
        orientation: 'portrait',
        margins: { top: 18, right: 16, bottom: 18, left: 16 },
        scale: { type: 'fixed', value: 1 },
      },
      repeatRows: {
        sheetId: 'invoice-sheet' as never,
        start: { row: 0, column: 0 },
        end: { row: 4, column: 5 },
      },
      manualBreaks: [{ sheetId: 'invoice-sheet' as never, beforeRow: 14 }],
      header: { right: 'Invoice {{page}}' },
      footer: { center: 'Generated by tego-sheet · {{page}}/{{pages}}' },
      showGridlines: false,
      showHeadings: false,
    },
  ],
};
```

The fixture data is:

```ts
const data = Object.freeze({
  customer: Object.freeze({ name: 'Acme GmbH', address: 'Berlin' }),
  invoice: Object.freeze({ id: 'INV-2026-042', currency: 'EUR' }),
  items: Object.freeze([
    Object.freeze({ description: 'Hosting', quantity: 1, amount: 29 }),
    Object.freeze({ description: 'Support', quantity: 4, amount: 75 }),
    Object.freeze({ description: 'Training', quantity: 2, amount: 240 }),
  ]),
});
```

The render environment uses `clock: new Date('2026-07-25T00:00:00.000Z')`, locale `en-US`, time
zone `UTC`, Excel 1900 date system, and deterministic Arial font metrics.

- [ ] **Step 4: Implement the abortable pipeline**

```ts
export async function renderOutputRevision(
  request: OutputRevisionRequest,
): Promise<OutputRevisionResult> {
  const compilation = compileSpreadsheetTemplate(request.document, request.template);
  if (compilation.template === undefined) {
    return Object.freeze({ revision: request.revision, diagnostics: compilation.diagnostics });
  }
  const rendered = await renderSpreadsheetTemplate(
    {
      template: compilation.template,
      currentDocumentHash: hashSpreadsheetDocument(request.document),
      data: request.data,
      profileId: request.template.printProfiles[0]!.id,
      missingValue: 'error',
      signal: request.signal,
    },
    request.environment,
  );
  return Object.freeze({
    revision: request.revision,
    diagnostics: Object.freeze([...compilation.diagnostics, ...rendered.diagnostics]),
    ...(rendered.document === undefined ? {} : { document: rendered.document }),
  });
}
```

- [ ] **Step 5: Run pipeline and template tests**

Run:

```bash
npx vitest run --project unit tests/unit/website/output-studio-pipeline.test.ts tests/unit/template/render.test.ts
```

Expected: PASS with the fixture producing exactly two pages.

- [ ] **Step 6: Commit**

```bash
git add website/src/components/playground/output-studio-fixtures.ts \
  website/src/components/playground/output-studio-pipeline.ts \
  tests/unit/website/output-studio-pipeline.test.ts
git commit -m "feat(docs): add deterministic output fixture"
```

---

### Task 4: Output Studio preview and restricted template workbench

**Files:**

- Create: `website/src/components/playground/output-studio.tsx`
- Create: `tests/component/docs-output-studio.test.tsx`
- Modify: `website/src/components/playground/playground.module.css`

**Interfaces:**

- Produces:
  - `OutputStudio`
  - named regions `Output inputs`, `Exact page preview`, `Pipeline and outputs`
- Consumes:
  - Task 2 state model
  - Task 3 fixture and pipeline
  - public `TegoSheet`, `TemplateDesigner`, and `TemplatePreview`

- [ ] **Step 1: Write the failing initial-preview component test**

Mock only the Task 3 pipeline result. Keep `TemplatePreview` real so page SVG serialization remains
covered.

```tsx
it('renders the prepared revision and explains the shared artifact', async () => {
  render(<OutputStudio />);

  expect(screen.getByRole('heading', { name: 'Output Studio' })).toBeTruthy();
  await waitFor(() => expect(screen.getAllByRole('article', { name: /Print page/ })).toHaveLength(2));
  expect(screen.getByText('GeneratedDocument · revision 1')).toBeTruthy();
  expect(screen.getByText(/One document · many outputs/)).toBeTruthy();
});
```

- [ ] **Step 2: Run the component test and verify RED**

Run:

```bash
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Expected: FAIL because `OutputStudio` does not exist.

- [ ] **Step 3: Implement initial rendering and exact preview**

`OutputStudio` initializes the fixture once, starts revision 1 in an effect, aborts on cleanup, and
commits only matching results:

```tsx
const [fixture] = useState(createInvoiceOutputFixture);
const [draftTemplate, setDraftTemplate] = useState(fixture.template);
const [draftData, setDraftData] = useState(() => JSON.stringify(fixture.data, null, 2));
const [state, dispatch] = useReducer(reduceOutputStudioState, undefined, createOutputStudioState);

useEffect(() => {
  const controller = new AbortController();
  const revision = 1;
  dispatch({ type: 'render-started', revision });
  void renderOutputRevision({
    revision,
    document: fixture.document,
    template: fixture.template,
    data: fixture.data,
    environment: fixture.environment,
    signal: controller.signal,
  }).then((result) => {
    if (controller.signal.aborted) return;
    if (result.document === undefined)
      dispatch({ type: 'render-blocked', revision, diagnostics: result.diagnostics });
    else
      dispatch({
        type: 'render-succeeded',
        revision,
        document: result.document,
        diagnostics: result.diagnostics,
      });
  });
  return () => controller.abort();
}, [fixture]);
```

Render `<TemplatePreview document={state.generatedDocument} />` only when a successful document
exists.

- [ ] **Step 4: Add a failing stale-draft and workbench test**

```tsx
it('keeps template edits as drafts until Apply and regenerate', async () => {
  render(<OutputStudio />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Edit template' }));
  fireEvent.change(screen.getByLabelText('Expression for customer-name'), {
    target: { value: 'customer.legalName' },
  });

  expect(screen.getByRole('status').textContent).toContain('Preview is stale');
  expect((screen.getByRole('button', { name: 'Print 2 pages' }) as HTMLButtonElement).disabled).toBe(
    true,
  );
  expect(document.querySelector('[data-mode="template"]')).toBeTruthy();
});
```

- [ ] **Step 5: Implement restricted workbench and explicit regeneration**

- `Edit template` expands a named workbench region.
- Render `TegoSheet` with `document`, `mode="template"`, `template`, and `onTemplateChange`.
- Render `TemplateDesigner` beside the sheet only through the public template-mode surface already
  provided by `TegoSheet`; do not duplicate its editor controls.
- JSON data uses a controlled textarea.
- Template or data edits dispatch `draft-changed`.
- `Apply & regenerate` parses JSON, increments the revision, aborts the prior controller, and invokes
  `renderOutputRevision`.

- [ ] **Step 6: Run component and template-designer tests**

Run:

```bash
npx vitest run --project component tests/component/docs-output-studio.test.tsx tests/component/template-designer.test.tsx
```

Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add website/src/components/playground/output-studio.tsx \
  website/src/components/playground/playground.module.css \
  tests/component/docs-output-studio.test.tsx
git commit -m "feat(docs): add output studio preview"
```

---

### Task 5: Output adapters and browser download lifecycle

**Files:**

- Create: `website/src/components/playground/output-download.ts`
- Modify: `website/src/components/playground/output-studio.tsx`
- Modify: `tests/component/docs-output-studio.test.tsx`

**Interfaces:**

- Produces:
  - `downloadBlob(blob, filename, document, url): void`
  - injectable `OutputStudioAdapters`
- Consumes:
  - `IsolatedBrowserPrintAdapter`
  - `PdfAdapter` from `tego-sheet/output/pdf`
  - `ImageAdapter` from `tego-sheet/output/image`
  - `XlsxAdapter` from `tego-sheet/output/xlsx`

- [ ] **Step 1: Write failing same-document adapter tests**

```tsx
it('passes the current generated document to every output adapter', async () => {
  const adapters = createAdapterDoubles();
  render(<OutputStudio adapters={adapters} />);
  await screen.findByText('GeneratedDocument · revision 1');

  fireEvent.click(screen.getByRole('button', { name: 'Print 2 pages' }));
  await waitFor(() => expect(adapters.print.print).toHaveBeenCalledOnce());

  fireEvent.click(screen.getByRole('button', { name: 'Download PDF' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download PNG page 1' }));
  fireEvent.click(screen.getByRole('button', { name: 'Download XLSX' }));

  await waitFor(() => expect(adapters.xlsx.render).toHaveBeenCalledOnce());
  const generated = adapters.print.print.mock.calls[0]![0];
  expect(adapters.pdf.render.mock.calls[0]![0]).toBe(generated);
  expect(adapters.image.render.mock.calls[0]![0]).toBe(generated);
  expect(adapters.xlsx.render.mock.calls[0]![0]).toBe(generated);
});
```

- [ ] **Step 2: Run the adapter test and verify RED**

Run:

```bash
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Expected: FAIL because output actions and adapter injection are absent.

- [ ] **Step 3: Implement Blob download cleanup**

```ts
export function downloadBlob(
  blob: Blob,
  filename: string,
  targetDocument: Document = document,
  url: Pick<typeof URL, 'createObjectURL' | 'revokeObjectURL'> = URL,
): void {
  const href = url.createObjectURL(blob);
  const anchor = targetDocument.createElement('a');
  anchor.href = href;
  anchor.download = filename;
  anchor.hidden = true;
  targetDocument.body.append(anchor);
  try {
    anchor.click();
  } finally {
    anchor.remove();
    queueMicrotask(() => url.revokeObjectURL(href));
  }
}
```

- [ ] **Step 4: Implement injectable adapter actions**

```ts
export interface OutputStudioAdapters {
  readonly print: Pick<IsolatedBrowserPrintAdapter, 'print' | 'dispose'>;
  readonly pdf: Pick<PdfAdapter, 'render'>;
  readonly image: Pick<ImageAdapter, 'render'>;
  readonly xlsx: Pick<XlsxAdapter, 'render'>;
}
```

Default adapters are created once and disposed on unmount. Actions use:

```ts
await adapters.print.print(generated);
const pdf = await adapters.pdf.render(generated, {
  pages: 'all',
  metadata: { title },
  tagged: false,
});
const [png] = await adapters.image.render(generated, {
  format: 'png',
  pages: [selectedPage],
  background: '#ffffff',
  dpi: 144,
});
const xlsx = await adapters.xlsx.render(generated, {
  formulaMode: 'formula-and-cached-value',
  compatibility: 'excel',
});
```

Every handler dispatches `output-started`, then `output-finished` or `output-failed`. Output buttons
are disabled unless phase is `ready`.

- [ ] **Step 5: Add and run cleanup/failure tests**

Test exact MIME types, filenames, URL revocation, independent busy state, `PRINT_BLOCKED`, and an
XLSX rejection that leaves the preview visible.

Run:

```bash
npx vitest run --project component tests/component/docs-output-studio.test.tsx
```

Expected: PASS with no unhandled promise rejection.

- [ ] **Step 6: Commit**

```bash
git add website/src/components/playground/output-download.ts \
  website/src/components/playground/output-studio.tsx \
  tests/component/docs-output-studio.test.tsx
git commit -m "feat(docs): connect output studio adapters"
```

---

### Task 6: Integrate the workspace shell without regressing Spreadsheet

**Files:**

- Modify: `website/src/components/playground/playground.tsx`
- Modify: `website/src/components/playground/playground.module.css`
- Modify: `tests/component/docs-playground.test.tsx`
- Modify: `tests/architecture/documentation-site-contract.test.ts`

**Interfaces:**

- Consumes: `OutputStudio`, Task 1 workspace URL helpers.
- Produces: top-level accessible workspace navigation and compatible Spreadsheet behavior.

- [ ] **Step 1: Add failing shell compatibility tests**

```tsx
it('keeps Spreadsheet as the default for legacy mode URLs', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?mode=controlled');
  await renderPlayground();
  expect(screen.getByRole('tab', { name: 'Spreadsheet' }).getAttribute('aria-selected')).toBe(
    'true',
  );
  expect(screen.getByTestId('tego-sheet-double').getAttribute('data-ownership')).toBe('controlled');
});

it('opens Output Studio directly and restores Spreadsheet through history', async () => {
  window.history.replaceState({}, '', '/tego-sheet/playground?workspace=output&mode=locales');
  await renderPlayground();
  expect(screen.getByRole('heading', { name: 'Output Studio' })).toBeTruthy();
  fireEvent.click(screen.getByRole('tab', { name: 'Spreadsheet' }));
  expect((screen.getByRole('radio', { name: 'Locales' }) as HTMLInputElement).checked).toBe(true);
});
```

- [ ] **Step 2: Run Playground component tests and verify RED**

Run:

```bash
npx vitest run --project component tests/component/docs-playground.test.tsx
```

Expected: new tests FAIL because workspace tabs are absent.

- [ ] **Step 3: Extract `SpreadsheetWorkspace` and add the shell**

Keep existing `PresetSession` and inspector code unchanged inside `SpreadsheetWorkspace`.
`Playground` owns:

```tsx
const [location, setLocation] = useState(() =>
  readPlaygroundLocation(window.location.search),
);

return (
  <main className={styles.playground}>
    <header className={styles.workspaceHeader}>
      <div>
        <p className={styles.eyebrow}>Live public API examples</p>
        <h1>Playground</h1>
      </div>
      <div role="tablist" aria-label="Playground workspace">
        <button role="tab" aria-selected={location.workspace === 'spreadsheet'}>
          Spreadsheet
        </button>
        <button role="tab" aria-selected={location.workspace === 'output'}>
          Output Studio
        </button>
      </div>
    </header>
    {location.workspace === 'spreadsheet' ? (
      <SpreadsheetWorkspace mode={location.mode} />
    ) : (
      <OutputStudio />
    )}
  </main>
);
```

History selection uses `writePlaygroundLocation`; the popstate listener uses
`readPlaygroundLocation`.

- [ ] **Step 4: Add responsive layout CSS**

Use named classes rather than structural selectors:

```css
.outputStudioGrid {
  display: grid;
  grid-template-columns: minmax(18rem, 0.8fr) minmax(28rem, 1.6fr) minmax(18rem, 0.85fr);
  gap: 1rem;
}

@media (max-width: 72rem) {
  .outputStudioGrid {
    grid-template-columns: 1fr 1fr;
  }
  .outputPreview {
    grid-column: 1 / -1;
    grid-row: 1;
  }
}

@media (max-width: 42rem) {
  .outputStudioGrid {
    display: flex;
    flex-direction: column;
  }
  .outputPreview {
    order: 1;
  }
  .outputInputs {
    order: 2;
  }
  .outputPipeline {
    order: 3;
  }
}
```

- [ ] **Step 5: Run component and architecture tests**

Run:

```bash
npx vitest run --project component tests/component/docs-playground.test.tsx tests/component/docs-output-studio.test.tsx
npx vitest run --project architecture tests/architecture/documentation-site-contract.test.ts
```

Expected: PASS, including the existing 44-pixel target contract.

- [ ] **Step 6: Commit**

```bash
git add website/src/components/playground/playground.tsx \
  website/src/components/playground/playground.module.css \
  tests/component/docs-playground.test.tsx \
  tests/architecture/documentation-site-contract.test.ts
git commit -m "feat(docs): add playground workspace shell"
```

---

### Task 7: Documentation browser and visual coverage

**Files:**

- Modify: `tests/docs/docs.spec.ts`
- Modify: `tests/docs-visual/docs-visual.spec.ts`
- Modify: `website/docs/guides/printing.md`

**Interfaces:**

- Consumes: canonical Output Studio URL and accessible labels from Tasks 4-6.
- Produces: runnable documentation link and browser-level regression coverage.

- [ ] **Step 1: Write the failing docs browser test**

```ts
test('Output Studio regenerates one revision and exposes deterministic outputs', async ({
  page,
}) => {
  await page.goto('/tego-sheet/playground?workspace=output');
  await expect(page.getByRole('heading', { name: 'Output Studio' })).toBeVisible();
  await expect(page.getByText('GeneratedDocument · revision 1')).toBeVisible();
  await expect(page.getByRole('article', { name: /Print page/ })).toHaveCount(2);

  await page.getByRole('button', { name: 'Edit sample data' }).click();
  await page.getByLabel('Sample data JSON').fill(
    JSON.stringify({
      customer: { name: 'Northwind Traders', address: 'Berlin' },
      invoice: { id: 'INV-2026-043', currency: 'EUR' },
      items: [
        { description: 'Hosting', quantity: 1, amount: 29 },
        { description: 'Support', quantity: 4, amount: 75 },
        { description: 'Training', quantity: 2, amount: 240 },
      ],
    }),
  );
  await expect(page.getByRole('status')).toContainText('Preview is stale');
  await expect(page.getByRole('button', { name: 'Download PDF' })).toBeDisabled();
  await page.getByRole('button', { name: 'Apply & regenerate' }).click();
  await expect(page.getByText('GeneratedDocument · revision 2')).toBeVisible();
});
```

- [ ] **Step 2: Run the targeted docs test and verify RED**

Run:

```bash
npm run docs:build
npm run test:docs -- --grep "Output Studio"
```

Expected: FAIL until the built page contains the new workspace and exact labels.

- [ ] **Step 3: Complete docs behavior coverage**

Add tests for:

- direct canonical URL;
- keyboard workspace switching;
- invalid JSON blocking;
- stale output disabling;
- successful regeneration;
- output button labels without clicking native Print.

- [ ] **Step 4: Add visual cases**

Use the existing visual harness to capture:

- desktop `1440 × 1100` ready state;
- intermediate `1024 × 1100` ready state;
- narrow `390 × 1000` ready state;
- desktop stale state;
- desktop blocked diagnostic state.

- [ ] **Step 5: Link the demo from the printing guide**

Add directly below the guide introduction:

```md
Try the complete pipeline in
[Output Studio](/playground?workspace=output): edit the prepared invoice template and sample data,
regenerate its exact pages, then inspect browser print, PDF, PNG, and XLSX output actions.
```

- [ ] **Step 6: Run docs and visual tests**

Run:

```bash
npm run docs:build
npm run test:docs
npm run test:docs-visual -- --update-snapshots
npm run test:docs-visual
```

Expected: all docs tests and visual baselines PASS.

- [ ] **Step 7: Commit**

```bash
git add tests/docs/docs.spec.ts tests/docs-visual/docs-visual.spec.ts \
  website/docs/guides/printing.md tests/docs-visual
git commit -m "test(docs): cover output studio workflow"
```

---

### Task 8: Integrated verification and cleanup

**Files:**

- Modify only files required to fix verified failures from the commands below.

**Interfaces:**

- Consumes all prior tasks.
- Produces a clean, release-ready worktree with fresh evidence.

- [ ] **Step 1: Run targeted feature tests**

```bash
npx vitest run --project unit \
  tests/unit/website/playground-workspace.test.ts \
  tests/unit/website/output-studio-model.test.ts \
  tests/unit/website/output-studio-pipeline.test.ts
npx vitest run --project component \
  tests/component/docs-playground.test.tsx \
  tests/component/docs-output-studio.test.tsx \
  tests/component/template-designer.test.tsx
```

Expected: all targeted files PASS.

- [ ] **Step 2: Run static verification**

```bash
npm run format:check
npm run lint
npm run typecheck
npm run typecheck:docs
```

Expected: every command exits 0 with no warnings.

- [ ] **Step 3: Run package and documentation gates**

```bash
npm run build
npm run test:package
npm run docs:build
npm run test:docs
npm run test:docs-visual
```

Expected: every command exits 0.

- [ ] **Step 4: Run the full test suite**

```bash
npm test
npm run test:ssr
```

Expected: all non-environment-gated tests PASS.

- [ ] **Step 5: Perform browser visual QA**

Start:

```bash
npm run docs:serve-static
```

Inspect `/tego-sheet/playground?workspace=output` in the in-app browser at desktop and narrow widths.
Verify:

- no clipping or horizontal page overflow;
- exact two-page initial preview;
- workbench opens and closes;
- stale state disables all outputs;
- invalid JSON is localized and recoverable;
- regeneration replaces the revision atomically;
- every control has a visible focus state;
- system Print is not invoked during QA.

- [ ] **Step 6: Commit verification-only fixes**

If verification required changes:

```bash
git add website/src/components/playground tests/unit/website tests/component \
  tests/docs tests/docs-visual website/docs/guides/printing.md
git commit -m "fix(docs): harden output studio verification"
```

If no files changed, do not create an empty commit.

- [ ] **Step 7: Confirm clean status**

```bash
git status --short --branch
```

Expected: no unstaged or staged files; branch contains only intentional commits.
