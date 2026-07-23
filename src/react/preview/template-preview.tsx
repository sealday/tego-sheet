import type { ReactElement } from 'react';
import {
  serializeGeneratedDocumentSvgPages,
  type GeneratedDocumentForBrowserPrint,
} from '../../output/browser-print-adapter';

export interface TemplatePreviewProps {
  readonly document: GeneratedDocumentForBrowserPrint;
}

/** Renders the exact SVG pages later submitted to the isolated browser adapter. */
export function TemplatePreview({ document }: TemplatePreviewProps): ReactElement {
  const pages = serializeGeneratedDocumentSvgPages(document);
  return (
    <section aria-label="Template print preview">
      {pages.map((page) => (
        <article
          key={page.id}
          aria-label={`Print page ${page.id}`}
          data-page-id={page.id}
          data-page-width={page.width}
          data-page-height={page.height}
          dangerouslySetInnerHTML={{ __html: page.svg }}
        />
      ))}
    </section>
  );
}
