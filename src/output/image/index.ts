import {
  serializeGeneratedDocumentSvgPages,
  type GeneratedDocumentForBrowserPrint,
} from '../browser-print-adapter';
import type { PrintDisplayCommand } from '../../print';
import type { GeneratedDocument, ResolvedResource } from '../../template';
import type { FontkitFont } from 'fontkit';
import {
  forbiddenFontEmbedding,
  forbiddenFontSubsetting,
  openTypeEmbeddingFlags,
} from '../font-embedding';
import { outputError, throwIfAborted } from '../output-error';

const DEFAULT_LIMITS: ImageOutputLimits = Object.freeze({
  maxPages: 1_000,
  maxPixels: 100_000_000,
  maxDpi: 600,
});

/** Image conversion request. */
export interface ImageOutputOptions {
  /** Vector SVG or raster PNG output. */
  readonly format: 'svg' | 'png';
  /** Ordered zero-based page selection. */
  readonly pages: readonly number[];
  /** PNG resolution; defaults to CSS 96 DPI. */
  readonly dpi?: number;
  /** Page background or transparent output. */
  readonly background?: string | 'transparent';
  /** Cancels conversion without returning a partial array. */
  readonly signal?: AbortSignal;
}

/** Bounded image output resources. */
export interface ImageOutputLimits {
  /** Maximum selected pages. */
  readonly maxPages: number;
  /** Maximum total raster pixels across selected pages. */
  readonly maxPixels: number;
  /** Maximum accepted PNG DPI. */
  readonly maxDpi: number;
}

/** One host-independent SVG rasterization request. */
export interface ImageRasterizeRequest {
  /** Safe standalone SVG source. */
  readonly svg: string;
  /** Exact output pixel width. */
  readonly width: number;
  /** Exact output pixel height. */
  readonly height: number;
  /** Source page width in device-independent units. */
  readonly pageWidth: number;
  /** Source page height in device-independent units. */
  readonly pageHeight: number;
  /** Immutable page display commands. */
  readonly commands: readonly PrintDisplayCommand[];
  /** Immutable resolved resources used by display commands. */
  readonly resources: GeneratedDocument['resources'];
  /** Requested output DPI. */
  readonly dpi: number;
  /** Explicit output background. */
  readonly background: string | 'transparent';
  /** Shared cancellation signal. */
  readonly signal?: AbortSignal;
}

/** Host rasterizer capability used by the PNG branch. */
export type ImageRasterizer = (request: ImageRasterizeRequest) => Promise<Blob>;

/** Image adapter construction options. */
export interface ImageAdapterOptions {
  /** Optional host PNG rasterizer; a browser default is used when available. */
  readonly rasterize?: ImageRasterizer;
  /** Optional bounded output limits. */
  readonly limits?: Partial<ImageOutputLimits>;
}

function escapeMarkup(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function selectedPages(document: GeneratedDocument, pages: readonly number[]): readonly number[] {
  const pageCount = document.print.pages.length;
  if (
    pages.length === 0 ||
    pages.some((page) => !Number.isInteger(page) || page < 0 || page >= pageCount)
  ) {
    throw outputError('IMAGE_PAGE_SELECTION_INVALID', 'Image page selection is invalid', {
      details: { pages, pageCount },
    });
  }
  return pages;
}

function safeBackground(value: string | undefined): string | 'transparent' {
  const background = value ?? 'transparent';
  if (
    background !== 'transparent' &&
    !/^(?:#[\da-f]{3,8}|[a-z]+|rgba?\([\d\s.,%+-]+\)|hsla?\([\d\s.,%+-]+\))$/iu.test(background)
  ) {
    throw outputError('IMAGE_ENCODING_FAILED', 'Image background is not a safe CSS color');
  }
  return background;
}

function isStandardFont(family: string): boolean {
  return /^(?:arial|helvetica|sans-serif|times(?: new roman)?|serif|courier|monospace)$/iu.test(
    family.trim(),
  );
}

function fontResources(document: GeneratedDocument): ReadonlyMap<string, ResolvedResource> {
  const resources = new Map<string, ResolvedResource>();
  for (const resource of Object.values(document.resources.byHash).sort((left, right) =>
    left.contentHash < right.contentHash ? -1 : left.contentHash > right.contentHash ? 1 : 0,
  )) {
    if (
      resource.type === 'font' &&
      resource.fontFamily !== undefined &&
      !resources.has(resource.fontFamily)
    ) {
      resources.set(resource.fontFamily, resource);
    }
  }
  return resources;
}

function assertOutlinePermission(resource: ResolvedResource): void {
  const flags = openTypeEmbeddingFlags(resource.bytes);
  if (flags === undefined || forbiddenFontEmbedding(flags) || forbiddenFontSubsetting(flags)) {
    throw outputError(
      'IMAGE_FONT_EMBEDDING_FAILED',
      `Font ${resource.fontFamily ?? resource.contentHash} forbids used-glyph output`,
      {
        details: {
          resource: resource.contentHash,
          ...(flags === undefined ? {} : { flags }),
        },
      },
    );
  }
}

async function outlineText(
  command: Extract<PrintDisplayCommand, { readonly kind: 'text' }>,
  resources: ReadonlyMap<string, ResolvedResource>,
  fonts: Map<string, FontkitFont>,
): Promise<PrintDisplayCommand> {
  if (isStandardFont(command.fontFamily)) return command;
  const resource = resources.get(command.fontFamily);
  if (resource === undefined) {
    throw outputError(
      'IMAGE_FONT_EMBEDDING_FAILED',
      `No resolved image font matches ${command.fontFamily}`,
      { details: { fontFamily: command.fontFamily } },
    );
  }
  assertOutlinePermission(resource);
  let font = fonts.get(resource.contentHash);
  if (font === undefined) {
    const { create } = await import('fontkit');
    font = create(new Uint8Array(resource.bytes));
    fonts.set(resource.contentHash, font);
  }
  for (const character of command.text) {
    if (!font.hasGlyphForCodePoint(character.codePointAt(0)!)) {
      throw outputError(
        'IMAGE_FONT_EMBEDDING_FAILED',
        `Font ${command.fontFamily} cannot represent selected text`,
        { details: { fontFamily: command.fontFamily } },
      );
    }
  }
  const run = font.layout(command.text);
  const scaleY = command.fontSize / font.unitsPerEm;
  const naturalWidth = run.advanceWidth * scaleY;
  const widthScale =
    naturalWidth > command.maxWidth && naturalWidth > 0 ? command.maxWidth / naturalWidth : 1;
  const scaleX = scaleY * widthScale;
  const renderedWidth = run.advanceWidth * scaleX;
  const startX =
    command.horizontalAlign === 'center'
      ? command.x - renderedWidth / 2
      : command.horizontalAlign === 'right'
        ? command.x - renderedWidth
        : command.x;
  let advance = 0;
  const paths = run.glyphs.map((glyph, index) => {
    const position = run.positions[index]!;
    const path = glyph.path
      .transform(
        scaleX,
        0,
        0,
        -scaleY,
        startX + (advance + position.xOffset) * scaleX,
        command.y - position.yOffset * scaleY,
      )
      .toSVG();
    advance += position.xAdvance;
    return path;
  });
  return Object.freeze({
    kind: 'path',
    data: paths.join(''),
    fill: command.color,
  });
}

async function imageCommands(
  document: GeneratedDocument,
  commands: readonly PrintDisplayCommand[],
  fonts: Map<string, FontkitFont>,
): Promise<readonly PrintDisplayCommand[]> {
  const resources = fontResources(document);
  const output: PrintDisplayCommand[] = [];
  for (const command of commands) {
    if (command.kind === 'text') {
      output.push(await outlineText(command, resources, fonts));
    } else if (command.kind === 'clip') {
      output.push(
        Object.freeze({
          ...command,
          commands: Object.freeze(await imageCommands(document, command.commands, fonts)),
        }),
      );
    } else if (command.kind === 'link' && /^(?:https?:|mailto:)/iu.test(command.href)) {
      output.push(Object.freeze({ ...command, href: 'unsafe:removed' }));
    } else {
      output.push(command);
    }
  }
  return Object.freeze(output);
}

function serializeImagePage(
  document: GeneratedDocument,
  pageIndex: number,
  commands: readonly PrintDisplayCommand[],
  background: string | 'transparent',
): {
  readonly page: { readonly width: number; readonly height: number };
  readonly display: { readonly commands: readonly PrintDisplayCommand[] };
  readonly svg: string;
} {
  const semantic = document.print.pages[pageIndex]!;
  const display = document.print.displayList.pages[pageIndex]!;
  const projection: GeneratedDocumentForBrowserPrint = {
    print: {
      pages: [{ ...semantic, index: 0 }],
      displayList: {
        diagnostics: document.print.displayList.diagnostics,
        pages: [{ ...display, index: 0, commands }],
      },
    },
    resources: document.resources,
  };
  const page = serializeGeneratedDocumentSvgPages(projection)[0]!;
  const pageBackground =
    background === 'transparent'
      ? ''
      : `<rect width="100%" height="100%" fill="${escapeMarkup(background)}"/>`;
  const output = page.svg.replace(' role="img">', ` role="img">${pageBackground}`);
  if (/<script|<foreignObject|\s(?:src|href|xlink:href)=["']https?:/iu.test(output)) {
    throw outputError('IMAGE_ENCODING_FAILED', 'SVG output contains active external content');
  }
  return { page, display: { commands }, svg: output };
}

async function rasterizeWithBrowser(request: ImageRasterizeRequest): Promise<Blob> {
  throwIfAborted(request.signal);
  const source = new Blob([request.svg], { type: 'image/svg+xml' });
  if (typeof document !== 'undefined') {
    const canvas = document.createElement('canvas');
    canvas.width = request.width;
    canvas.height = request.height;
    const context = canvas.getContext('2d');
    if (context === null) throw new Error('Canvas 2D context is unavailable');
    const url = URL.createObjectURL(source);
    try {
      const image = new Image();
      image.src = url;
      await image.decode();
      throwIfAborted(request.signal);
      if (request.background !== 'transparent') {
        context.fillStyle = request.background;
        context.fillRect(0, 0, request.width, request.height);
      }
      context.drawImage(image, 0, 0, request.width, request.height);
      const blob = await new Promise<Blob>((resolve, reject) =>
        canvas.toBlob(
          (value) =>
            value === null ? reject(new Error('Canvas PNG encoding failed')) : resolve(value),
          'image/png',
        ),
      );
      throwIfAborted(request.signal);
      return blob;
    } finally {
      URL.revokeObjectURL(url);
      canvas.width = 0;
      canvas.height = 0;
    }
  }
  if (typeof globalThis.OffscreenCanvas === 'function') {
    return rasterizeDisplayList(request);
  }
  throw new Error('No SVG rasterizer is available in this environment');
}

type RasterContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

function imagePlacement(
  sourceWidth: number,
  sourceHeight: number,
  target: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  fit: 'contain' | 'cover' | 'fill',
): readonly [number, number, number, number] {
  if (fit === 'fill') return [target.x, target.y, target.width, target.height];
  const scale =
    fit === 'contain'
      ? Math.min(target.width / sourceWidth, target.height / sourceHeight)
      : Math.max(target.width / sourceWidth, target.height / sourceHeight);
  const width = sourceWidth * scale;
  const height = sourceHeight * scale;
  return [
    target.x + (target.width - width) / 2,
    target.y + (target.height - height) / 2,
    width,
    height,
  ];
}

async function paintCommands(
  context: RasterContext,
  commands: readonly PrintDisplayCommand[],
  request: ImageRasterizeRequest,
): Promise<void> {
  for (const command of commands) {
    throwIfAborted(request.signal);
    switch (command.kind) {
      case 'fill-rect':
        context.fillStyle = command.color;
        context.fillRect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
        break;
      case 'stroke-rect':
        context.strokeStyle = command.color;
        context.lineWidth = command.width;
        context.strokeRect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
        break;
      case 'line':
        context.beginPath();
        context.moveTo(command.x1, command.y1);
        context.lineTo(command.x2, command.y2);
        context.strokeStyle = command.color;
        context.lineWidth = command.width;
        context.stroke();
        break;
      case 'text':
        context.fillStyle = command.color;
        context.font = `${command.fontSize}px "${command.fontFamily}"`;
        context.textAlign = command.horizontalAlign;
        context.textBaseline = 'alphabetic';
        context.fillText(command.text, command.x, command.y, command.maxWidth);
        break;
      case 'path': {
        if (typeof globalThis.Path2D !== 'function') {
          throw new Error('Path2D is unavailable for Worker image rasterization');
        }
        const path = new Path2D(command.data);
        if (command.fill !== undefined) {
          context.fillStyle = command.fill;
          context.fill(path);
        }
        if (command.stroke !== undefined) {
          context.strokeStyle = command.stroke;
          context.lineWidth = command.width ?? 1;
          context.stroke(path);
        }
        break;
      }
      case 'clip':
        context.save();
        context.beginPath();
        context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
        context.clip();
        await paintCommands(context, command.commands, request);
        context.restore();
        break;
      case 'link':
        break;
      case 'image': {
        const resource = request.resources.byReference[command.resourceId];
        if (resource === undefined) {
          throw new Error(`Image resource ${command.resourceId} is missing`);
        }
        if (resource.vector !== undefined) {
          if (typeof globalThis.Path2D !== 'function') {
            throw new Error('Path2D is unavailable for Worker vector rasterization');
          }
          const vector = resource.vector;
          const [x, y, width, height] = imagePlacement(
            vector.viewBox[2],
            vector.viewBox[3],
            command.rect,
            command.fit,
          );
          context.save();
          context.beginPath();
          context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
          context.clip();
          context.translate(x, y);
          context.scale(width / vector.viewBox[2], height / vector.viewBox[3]);
          context.translate(-vector.viewBox[0], -vector.viewBox[1]);
          if (vector.background !== 'transparent') {
            context.fillStyle = vector.background;
            context.fillRect(...vector.viewBox);
          }
          context.fillStyle = vector.foreground;
          for (const data of vector.paths) context.fill(new Path2D(data));
          context.restore();
        } else if (resource.type === 'image') {
          if (typeof globalThis.createImageBitmap !== 'function') {
            throw new Error('createImageBitmap is unavailable for Worker image rasterization');
          }
          const bitmap = await createImageBitmap(
            new Blob([new Uint8Array(resource.bytes)], { type: resource.mimeType }),
          );
          try {
            const placement = imagePlacement(
              bitmap.width,
              bitmap.height,
              command.rect,
              command.fit,
            );
            context.save();
            context.beginPath();
            context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
            context.clip();
            context.drawImage(bitmap, ...placement);
            context.restore();
          } finally {
            bitmap.close();
          }
        } else {
          throw new Error(`Resource ${command.resourceId} is not an image`);
        }
        break;
      }
    }
  }
}

async function rasterizeDisplayList(request: ImageRasterizeRequest): Promise<Blob> {
  throwIfAborted(request.signal);
  const canvas = new OffscreenCanvas(request.width, request.height);
  const context = canvas.getContext('2d');
  if (context === null) throw new Error('Offscreen 2D context is unavailable');
  if (request.background !== 'transparent') {
    context.fillStyle = request.background;
    context.fillRect(0, 0, request.width, request.height);
  }
  context.scale(request.width / request.pageWidth, request.height / request.pageHeight);
  await paintCommands(context, request.commands, request);
  throwIfAborted(request.signal);
  return canvas.convertToBlob({ type: 'image/png' });
}

/** Pure SVG and PNG translation of immutable generated pages. */
export class ImageAdapter {
  readonly #limits: ImageOutputLimits;
  readonly #rasterize: ImageRasterizer;

  /** Creates a bounded image adapter. */
  constructor(options: ImageAdapterOptions = {}) {
    this.#limits = Object.freeze({ ...DEFAULT_LIMITS, ...options.limits });
    this.#rasterize = options.rasterize ?? rasterizeWithBrowser;
  }

  /** Renders selected pages atomically in request order. */
  async render(document: GeneratedDocument, options: ImageOutputOptions): Promise<readonly Blob[]> {
    throwIfAborted(options.signal);
    const pages = selectedPages(document, options.pages);
    if (pages.length > this.#limits.maxPages) {
      throw outputError('IMAGE_PIXEL_LIMIT_EXCEEDED', 'Image page count exceeds its limit');
    }
    const background = safeBackground(options.background);
    const dpi = options.dpi ?? 96;
    if (
      options.format === 'png' &&
      (!Number.isFinite(dpi) || dpi <= 0 || dpi > this.#limits.maxDpi)
    ) {
      throw outputError('IMAGE_ENCODING_FAILED', 'PNG DPI is outside the supported range');
    }
    const pixelSizes = pages.map((pageIndex) => ({
      width: Math.round((document.print.pages[pageIndex]!.width * dpi) / 96),
      height: Math.round((document.print.pages[pageIndex]!.height * dpi) / 96),
    }));
    const totalPixels = pixelSizes.reduce((sum, { width, height }) => sum + width * height, 0);
    if (
      options.format === 'png' &&
      (!Number.isSafeInteger(totalPixels) ||
        pixelSizes.some(({ width, height }) => width <= 0 || height <= 0) ||
        totalPixels > this.#limits.maxPixels)
    ) {
      throw outputError('IMAGE_PIXEL_LIMIT_EXCEEDED', 'PNG output exceeds its total pixel limit', {
        details: { dpi, totalPixels, maxPixels: this.#limits.maxPixels },
      });
    }
    const fonts = new Map<string, FontkitFont>();
    const selected = [];
    for (const pageIndex of pages) {
      throwIfAborted(options.signal);
      const display = document.print.displayList.pages[pageIndex];
      if (display === undefined) {
        throw outputError('IMAGE_ENCODING_FAILED', 'Image display page is missing');
      }
      const commands = await imageCommands(document, display.commands, fonts);
      selected.push(serializeImagePage(document, pageIndex, commands, background));
    }
    if (options.format === 'svg') {
      throwIfAborted(options.signal);
      return Object.freeze(selected.map(({ svg }) => new Blob([svg], { type: 'image/svg+xml' })));
    }
    const pixelPages = selected.map((entry, index) => ({
      ...entry,
      ...pixelSizes[index]!,
    }));
    const completed: Blob[] = [];
    try {
      for (const page of pixelPages) {
        throwIfAborted(options.signal);
        const encoded = await this.#rasterize({
          svg: page.svg,
          width: page.width,
          height: page.height,
          pageWidth: page.page.width,
          pageHeight: page.page.height,
          commands: page.display.commands,
          resources: document.resources,
          dpi,
          background,
          ...(options.signal === undefined ? {} : { signal: options.signal }),
        });
        throwIfAborted(options.signal);
        const buffer = await encoded.arrayBuffer();
        completed.push(new Blob([buffer], { type: 'image/png' }));
      }
      return Object.freeze(completed);
    } catch (cause) {
      if (cause instanceof Error && 'code' in cause) throw cause;
      throw outputError('IMAGE_ENCODING_FAILED', 'PNG encoding failed', { cause });
    }
  }
}

export { OutputAdapterError } from '../output-error';
