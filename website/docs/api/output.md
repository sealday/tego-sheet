---
sidebar_label: Output adapters
---

# Output adapters

Output adapters translate an already generated document; they do not recalculate formulas,
resolve template data, or paginate again.

```ts
import { ImageAdapter } from 'tego-sheet/output/image';
import { PdfAdapter } from 'tego-sheet/output/pdf';
import { XlsxAdapter } from 'tego-sheet/output/xlsx';

const pdf = await new PdfAdapter().render(generated, { signal });
const svgPages = await new ImageAdapter().render(generated, { format: 'svg', signal });
const xlsx = await new XlsxAdapter().render(generated, { signal });
```

All adapters are bounded and abortable. PDF preserves searchable vector text and checks font
embedding rights. Image output supports standalone SVG and DPI-bounded PNG. XLSX writes semantic
workbook data, print profiles, conditional formats, and anchored drawings with deterministic ZIP
metadata.

The XLSX implementation passes local structure and round-trip tests, but the roadmap remains
planned until recorded Excel Desktop, Excel for web, and LibreOffice interoperability checks pass.
Unsupported or degraded features produce stable diagnostics instead of silent loss.
