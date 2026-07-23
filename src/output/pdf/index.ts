import type PDFKit from 'pdfkit';
import type { PrintDisplayCommand } from '../../print';
import type { GeneratedDocument, ResolvedResource } from '../../template';
import {
  forbiddenFontEmbedding,
  forbiddenFontSubsetting,
  openTypeEmbeddingFlags,
} from '../font-embedding';
import { outputError, throwIfAborted } from '../output-error';

const POINTS_PER_CSS_PIXEL = 72 / 96;
const DEFAULT_LIMITS: PdfOutputLimits = Object.freeze({
  maxPages: 1_000,
  maxResourceBytes: 64 * 1024 * 1024,
  maxOutputBytes: 128 * 1024 * 1024,
  maxDurationMs: 30_000,
});

/** Metadata written into a generated PDF. */
export interface PdfMetadata {
  /** Document title. */
  readonly title?: string;
  /** Document author. */
  readonly author?: string;
  /** Document subject. */
  readonly subject?: string;
  /** Document creator. */
  readonly creator?: string;
  /** Search keywords. */
  readonly keywords?: readonly string[];
}

/** Bounded resources for one PDF conversion. */
export interface PdfOutputLimits {
  /** Maximum selected pages. */
  readonly maxPages: number;
  /** Maximum bytes read from resolved resources. */
  readonly maxResourceBytes: number;
  /** Maximum finalized PDF bytes. */
  readonly maxOutputBytes: number;
  /** Maximum elapsed conversion time. */
  readonly maxDurationMs: number;
}

/** PDF conversion request. */
export interface PdfOutputOptions {
  /** All pages or explicit zero-based page indices. */
  readonly pages: 'all' | readonly number[];
  /** Sanitized document metadata. */
  readonly metadata?: PdfMetadata;
  /** TP4 does not yet publish tagged-PDF semantics. */
  readonly tagged: false;
  /** Cancels conversion without returning a partial Blob. */
  readonly signal?: AbortSignal;
}

/** PDF adapter construction options. */
export interface PdfAdapterOptions {
  /** Optional bounded output limits. */
  readonly limits?: Partial<PdfOutputLimits>;
}

type PdfDocument = InstanceType<typeof PDFKit>;
type DestroyablePdfDocument = PdfDocument & {
  destroy(error?: Error): void;
};

function points(value: number): number {
  return value * POINTS_PER_CSS_PIXEL;
}

function sanitizeMetadata(value: string): string {
  return [...value]
    .filter((character) => {
      const codePoint = character.codePointAt(0)!;
      return codePoint >= 32 && codePoint !== 127;
    })
    .join('')
    .slice(0, 4_096);
}

function selectedPages(
  document: GeneratedDocument,
  selection: PdfOutputOptions['pages'],
): number[] {
  const count = document.print.pages.length;
  const pages =
    selection === 'all' ? Array.from({ length: count }, (_, index) => index) : [...selection];
  if (
    pages.length === 0 ||
    new Set(pages).size !== pages.length ||
    pages.some((page) => !Number.isInteger(page) || page < 0 || page >= count)
  ) {
    throw outputError('PDF_PAGE_SELECTION_INVALID', 'PDF page selection is invalid', {
      details: { pages, pageCount: count },
    });
  }
  return pages;
}

function isStandardFont(family: string): boolean {
  return /^(?:arial|helvetica|sans-serif|times(?: new roman)?|serif|courier|monospace)$/iu.test(
    family.trim(),
  );
}

function standardFont(family: string): string {
  const normalized = family.trim().toLowerCase();
  if (normalized.includes('times') || normalized === 'serif') return 'Times-Roman';
  if (normalized.includes('courier') || normalized === 'monospace') return 'Courier';
  return 'Helvetica';
}

function requiresEmbeddedFont(text: string): boolean {
  return [...text].some((character) => character.codePointAt(0)! > 0xff);
}

function fontResources(document: GeneratedDocument): ReadonlyMap<string, ResolvedResource> {
  const entries = Object.values(document.resources.byHash)
    .filter(
      (resource): resource is ResolvedResource & { readonly fontFamily: string } =>
        resource.type === 'font' && resource.fontFamily !== undefined,
    )
    .sort((left, right) => left.contentHash.localeCompare(right.contentHash));
  const resources = new Map<string, ResolvedResource>();
  for (const resource of entries) {
    if (!resources.has(resource.fontFamily)) resources.set(resource.fontFamily, resource);
  }
  return resources;
}

function checkFontPermission(resource: ResolvedResource): void {
  const fsType = openTypeEmbeddingFlags(resource.bytes);
  if (fsType === undefined) {
    throw outputError(
      'PDF_FONT_SUBSET_FAILED',
      `Font ${resource.fontFamily ?? resource.contentHash} has no readable OS/2 embedding policy`,
      { details: { resource: resource.contentHash } },
    );
  }
  if (forbiddenFontEmbedding(fsType)) {
    throw outputError(
      'PDF_FONT_EMBEDDING_FORBIDDEN',
      `Font ${resource.fontFamily ?? resource.contentHash} forbids embedding`,
      { details: { resource: resource.contentHash, fsType } },
    );
  }
  if (forbiddenFontSubsetting(fsType)) {
    throw outputError(
      'PDF_FONT_SUBSET_FAILED',
      `Font ${resource.fontFamily ?? resource.contentHash} forbids subsetting`,
      { details: { resource: resource.contentHash, fsType } },
    );
  }
}

function safeLink(href: string): boolean {
  return href.startsWith('#') || /^(?:https?:|mailto:)/iu.test(href);
}

function textBox(command: Extract<PrintDisplayCommand, { readonly kind: 'text' }>): {
  readonly x: number;
  readonly width: number;
} {
  if (command.horizontalAlign === 'center') {
    return { x: command.x - command.maxWidth / 2, width: command.maxWidth };
  }
  if (command.horizontalAlign === 'right') {
    return { x: command.x - command.maxWidth, width: command.maxWidth };
  }
  return { x: command.x, width: command.maxWidth };
}

function drawVectorResource(
  pdf: PdfDocument,
  command: Extract<PrintDisplayCommand, { readonly kind: 'image' }>,
  resource: ResolvedResource,
): void {
  const vector = resource.vector;
  if (vector === undefined) return;
  const [viewX, viewY, viewWidth, viewHeight] = vector.viewBox;
  const scaleX = command.rect.width / viewWidth;
  const scaleY = command.rect.height / viewHeight;
  const scale =
    command.fit === 'fill'
      ? undefined
      : command.fit === 'cover'
        ? Math.max(scaleX, scaleY)
        : Math.min(scaleX, scaleY);
  const renderedWidth = viewWidth * (scale ?? scaleX);
  const renderedHeight = viewHeight * (scale ?? scaleY);
  const offsetX = command.rect.x + (command.rect.width - renderedWidth) / 2;
  const offsetY = command.rect.y + (command.rect.height - renderedHeight) / 2;
  pdf.save();
  pdf.rect(
    points(command.rect.x),
    points(command.rect.y),
    points(command.rect.width),
    points(command.rect.height),
  );
  pdf.clip();
  pdf.translate(points(offsetX), points(offsetY));
  pdf.scale(POINTS_PER_CSS_PIXEL * (scale ?? scaleX), POINTS_PER_CSS_PIXEL * (scale ?? scaleY));
  pdf.translate(-viewX, -viewY);
  if (vector.background !== 'transparent') {
    pdf.rect(viewX, viewY, viewWidth, viewHeight).fill(vector.background);
  }
  for (const path of vector.paths) pdf.path(path).fill(vector.foreground);
  pdf.restore();
}

function drawCommands(
  pdf: PdfDocument,
  commands: readonly PrintDisplayCommand[],
  document: GeneratedDocument,
  fonts: ReadonlyMap<string, ResolvedResource>,
  deadline: number,
  signal?: AbortSignal,
): void {
  for (const candidate of commands as readonly unknown[]) {
    throwIfAborted(signal);
    if (Date.now() >= deadline) {
      throw outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF generation exceeded its time limit');
    }
    if (candidate === null || typeof candidate !== 'object' || Array.isArray(candidate)) {
      throw outputError('PDF_UNSUPPORTED_DRAW_COMMAND', 'PDF draw command is invalid');
    }
    const command = candidate as PrintDisplayCommand;
    switch (command.kind) {
      case 'fill-rect':
        pdf
          .rect(
            points(command.rect.x),
            points(command.rect.y),
            points(command.rect.width),
            points(command.rect.height),
          )
          .fill(command.color);
        break;
      case 'stroke-rect':
        pdf
          .lineWidth(points(command.width))
          .strokeColor(command.color)
          .rect(
            points(command.rect.x),
            points(command.rect.y),
            points(command.rect.width),
            points(command.rect.height),
          )
          .stroke();
        break;
      case 'line':
        pdf
          .lineWidth(points(command.width))
          .strokeColor(command.color)
          .moveTo(points(command.x1), points(command.y1))
          .lineTo(points(command.x2), points(command.y2))
          .stroke();
        break;
      case 'text': {
        const resource = fonts.get(command.fontFamily);
        if (requiresEmbeddedFont(command.text) && resource === undefined) {
          throw outputError(
            'PDF_FONT_SUBSET_FAILED',
            `No resolved font can encode ${command.fontFamily}`,
            { details: { fontFamily: command.fontFamily } },
          );
        }
        if (resource !== undefined) {
          checkFontPermission(resource);
          pdf.font(`tego:${resource.contentHash}`);
        } else if (isStandardFont(command.fontFamily)) {
          pdf.font(standardFont(command.fontFamily));
        } else {
          throw outputError(
            'PDF_FONT_SUBSET_FAILED',
            `No resolved font matches ${command.fontFamily}`,
            { details: { fontFamily: command.fontFamily } },
          );
        }
        const box = textBox(command);
        pdf
          .fontSize(points(command.fontSize))
          .fillColor(command.color)
          .text(command.text, points(box.x), points(command.y), {
            width: points(box.width),
            align: command.horizontalAlign,
            baseline: 'alphabetic',
            lineBreak: false,
          });
        break;
      }
      case 'image': {
        const resource = document.resources.byReference[command.resourceId];
        if (resource === undefined) {
          throw outputError('PDF_UNSUPPORTED_DRAW_COMMAND', 'PDF image resource is missing', {
            details: { resourceId: command.resourceId },
          });
        }
        if (resource.vector !== undefined) {
          drawVectorResource(pdf, command, resource);
          break;
        }
        if (resource.mimeType !== 'image/png' && resource.mimeType !== 'image/jpeg') {
          throw outputError('PDF_UNSUPPORTED_DRAW_COMMAND', 'PDF image type is unsupported', {
            details: { mimeType: resource.mimeType },
          });
        }
        const image = new Uint8Array(resource.bytes);
        const rectangle = command.rect;
        const size = [points(rectangle.width), points(rectangle.height)] as [number, number];
        const imageOptions: PDFKit.Mixins.ImageOption =
          command.fit === 'fill'
            ? { width: size[0], height: size[1] }
            : command.fit === 'cover'
              ? { cover: size, align: 'center', valign: 'center' }
              : { fit: size, align: 'center', valign: 'center' };
        pdf.image(
          image as unknown as PDFKit.Mixins.ImageSrc,
          points(rectangle.x),
          points(rectangle.y),
          imageOptions,
        );
        break;
      }
      case 'path':
        pdf.save();
        pdf.scale(POINTS_PER_CSS_PIXEL);
        pdf.path(command.data);
        if (command.width !== undefined) pdf.lineWidth(command.width);
        if (command.fill !== undefined && command.stroke !== undefined) {
          pdf.fillColor(command.fill).strokeColor(command.stroke).fillAndStroke();
        } else if (command.fill !== undefined) {
          pdf.fill(command.fill);
        } else if (command.stroke !== undefined) {
          pdf.stroke(command.stroke);
        }
        pdf.restore();
        break;
      case 'clip':
        pdf.save();
        pdf
          .rect(
            points(command.rect.x),
            points(command.rect.y),
            points(command.rect.width),
            points(command.rect.height),
          )
          .clip();
        drawCommands(pdf, command.commands, document, fonts, deadline, signal);
        pdf.restore();
        break;
      case 'link':
        if (safeLink(command.href)) {
          if (command.href.startsWith('#')) {
            pdf.goTo(
              points(command.rect.x),
              points(command.rect.y),
              points(command.rect.width),
              points(command.rect.height),
              command.href.slice(1),
            );
          } else {
            pdf.link(
              points(command.rect.x),
              points(command.rect.y),
              points(command.rect.width),
              points(command.rect.height),
              command.href,
            );
          }
        }
        break;
      default:
        throw outputError(
          'PDF_UNSUPPORTED_DRAW_COMMAND',
          `Unsupported PDF draw command ${(command as { readonly kind?: unknown }).kind as string}`,
        );
    }
  }
}

function collectPdf(
  pdf: DestroyablePdfDocument,
  signal: AbortSignal | undefined,
  maxOutputBytes: number,
  deadline: number,
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const chunks: BlobPart[] = [];
    let total = 0;
    let settled = false;
    const timeout = setTimeout(
      () => {
        pdf.destroy();
        fail(outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF generation exceeded its time limit'));
      },
      Math.max(0, deadline - Date.now()),
    );
    const cleanup = (): void => {
      clearTimeout(timeout);
      signal?.removeEventListener('abort', abort);
    };
    const fail = (error: unknown): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };
    const abort = (): void => {
      pdf.destroy();
      fail(outputError('RENDER_ABORTED', 'PDF generation was aborted'));
    };
    pdf.on('data', (chunk: Uint8Array) => {
      if (settled) return;
      const copy = new Uint8Array(chunk);
      total += copy.byteLength;
      if (total > maxOutputBytes) {
        pdf.destroy();
        fail(outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF output exceeds its byte limit'));
        return;
      }
      chunks.push(
        copy.buffer.slice(copy.byteOffset, copy.byteOffset + copy.byteLength) as ArrayBuffer,
      );
    });
    pdf.on('error', (cause: unknown) =>
      fail(outputError('PDF_FONT_SUBSET_FAILED', 'PDF generation failed', { cause })),
    );
    pdf.on('end', () => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(new Blob(chunks, { type: 'application/pdf' }));
    });
    signal?.addEventListener('abort', abort, { once: true });
  });
}

/** Pure PDF translation of an immutable generated document. */
export class PdfAdapter {
  readonly #limits: PdfOutputLimits;

  /** Creates a bounded PDF adapter. */
  constructor(options: PdfAdapterOptions = {}) {
    this.#limits = Object.freeze({ ...DEFAULT_LIMITS, ...options.limits });
  }

  /** Renders selected pre-paginated pages into one atomic PDF Blob. */
  async render(document: GeneratedDocument, options: PdfOutputOptions): Promise<Blob> {
    throwIfAborted(options.signal);
    const startedAt = Date.now();
    const deadline = startedAt + this.#limits.maxDurationMs;
    const pageIndices = selectedPages(document, options.pages);
    if (pageIndices.length > this.#limits.maxPages) {
      throw outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF page count exceeds its limit');
    }
    if (document.resources.totalBytes > this.#limits.maxResourceBytes) {
      throw outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF resources exceed their byte limit');
    }
    const module = await import('pdfkit/js/pdfkit.standalone.js');
    throwIfAborted(options.signal);
    if (Date.now() >= deadline) {
      throw outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF generation exceeded its time limit');
    }
    const PDFDocument = module.default;
    const info: PDFKit.PDFDocumentOptions['info'] = {
      Producer: 'tego-sheet',
      Creator: sanitizeMetadata(options.metadata?.creator ?? 'tego-sheet'),
      CreationDate: new Date(document.metadata.generatedAt),
      ModDate: new Date(document.metadata.generatedAt),
      ...(options.metadata?.title === undefined
        ? {}
        : { Title: sanitizeMetadata(options.metadata.title) }),
      ...(options.metadata?.author === undefined
        ? {}
        : { Author: sanitizeMetadata(options.metadata.author) }),
      ...(options.metadata?.subject === undefined
        ? {}
        : { Subject: sanitizeMetadata(options.metadata.subject) }),
      ...(options.metadata?.keywords === undefined
        ? {}
        : { Keywords: sanitizeMetadata(options.metadata.keywords.join(', ')) }),
    };
    const pdf = new PDFDocument({
      autoFirstPage: false,
      bufferPages: false,
      compress: false,
      info,
      tagged: false,
      fontLayoutCache: false,
    }) as DestroyablePdfDocument;
    const output = collectPdf(pdf, options.signal, this.#limits.maxOutputBytes, deadline);
    const fonts = fontResources(document);
    try {
      for (const resource of fonts.values()) {
        checkFontPermission(resource);
        pdf.registerFont(`tego:${resource.contentHash}`, new Uint8Array(resource.bytes));
      }
      for (const pageIndex of pageIndices) {
        throwIfAborted(options.signal);
        if (Date.now() >= deadline) {
          throw outputError('PDF_OUTPUT_LIMIT_EXCEEDED', 'PDF generation exceeded its time limit');
        }
        const semantic = document.print.pages[pageIndex]!;
        const display = document.print.displayList.pages[pageIndex];
        if (
          display === undefined ||
          display.index !== semantic.index ||
          display.width !== semantic.width ||
          display.height !== semantic.height
        ) {
          throw outputError(
            'PDF_UNSUPPORTED_DRAW_COMMAND',
            'Generated PDF pages and display list are inconsistent',
          );
        }
        pdf.addPage({
          size: [points(display.width), points(display.height)],
          margin: 0,
        });
        pdf.addNamedDestination(semantic.id);
        pdf.outline.addItem(semantic.id);
        drawCommands(pdf, display.commands, document, fonts, deadline, options.signal);
      }
      pdf.end();
      return await output;
    } catch (cause) {
      pdf.destroy();
      void output.catch(() => undefined);
      if (cause instanceof Error && 'code' in cause) throw cause;
      throw outputError('PDF_FONT_SUBSET_FAILED', 'PDF font embedding or subsetting failed', {
        cause,
      });
    }
  }
}

export { OutputAdapterError } from '../output-error';
