---
title: Template printing
---

# Template printing

Printing is a compiler pipeline. Compile an explicit template, render it against structured data and
deterministic environment inputs, then pass the immutable `GeneratedDocument` to preview or browser
output. The adapter never reads the editor DOM, selection, scroll position, zoom, or Canvas pixels.

Try the complete pipeline in
[Output Studio](/playground?workspace=output): edit the prepared invoice template and sample data,
regenerate its exact pages, then inspect browser print, PDF, PNG, and XLSX output actions. Preview,
browser print, PDF, and PNG consume the immutable page display list; XLSX instead writes the
semantic workbook carried by the same `GeneratedDocument`.

```ts
import {
  compileSpreadsheetTemplate,
  createFontMetrics,
  hashSpreadsheetDocument,
  IsolatedBrowserPrintAdapter,
  renderSpreadsheetTemplate,
  type SpreadsheetDocument,
  type SpreadsheetTemplate,
} from 'tego-sheet';

export async function printInvoice(
  document: SpreadsheetDocument,
  template: SpreadsheetTemplate,
  data: unknown,
) {
  const compilation = compileSpreadsheetTemplate(document, template);
  if (!compilation.template)
    throw new Error(compilation.diagnostics.map((item) => item.message).join('\n'));

  const rendered = await renderSpreadsheetTemplate(
    {
      template: compilation.template,
      currentDocumentHash: hashSpreadsheetDocument(document),
      data,
      profileId: template.printProfiles[0]!.id,
      missingValue: 'error',
    },
    {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: document.workbook.settings.dateSystem,
      clock: new Date(),
      fontMetrics: createFontMetrics({
        fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
        fallbackFont: 'Arial',
        fallback: { averageAdvance: 6, lineHeight: 12 },
      }),
    },
  );
  if (!rendered.document)
    throw new Error(rendered.diagnostics.map((item) => item.message).join('\n'));

  const adapter = new IsolatedBrowserPrintAdapter();
  try {
    await adapter.print(rendered.document);
  } finally {
    adapter.dispose();
  }
}
```

`TegoSheet` also supports `mode="template"` with a `template` and `onTemplateChange`, plus
`mode="preview"` with `sampleData`, `renderEnvironment`, and `onDiagnostics`. Preview and browser
print serialize the same page IDs, sizes, and `PrintDisplayList`.

An old compiled template is rejected with `TEMPLATE_SOURCE_STALE` when its canonical document hash
does not match the current snapshot. Browser printing uses a hidden same-origin iframe and removes it
after `afterprint`, cancellation, explicit disposal, or a bounded timeout.
