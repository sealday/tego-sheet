import type { DisplayRect, PrintDisplayCommand, PrintDisplayList } from '../print';

/** The immutable GeneratedDocument surface required by browser printing. */
export interface GeneratedDocumentForBrowserPrint {
  /** Pre-paginated output shared with preview adapters. */
  readonly print: {
    /** Compiler-issued semantic page identities. */
    readonly pages: readonly {
      /** Stable semantic identity shared with preview output. */
      readonly id: string;
      /** Zero-based page position in the generated document. */
      readonly index: number;
      /** Page width in print-profile units. */
      readonly width: number;
      /** Page height in print-profile units. */
      readonly height: number;
    }[];
    /** Exact display list shared with SVG preview and browser output. */
    readonly displayList: PrintDisplayList;
  };
  /** Ready immutable resources referenced by image display commands. */
  readonly resources?: {
    /** Logical resource references resolved for this render session. */
    readonly byReference: Readonly<
      Record<
        string,
        {
          /** Resolved logical resource category. */
          readonly type: string;
          /** Verified media type matching the stored bytes. */
          readonly mimeType: string;
          /** Immutable encoded resource payload. */
          readonly bytes: readonly number[];
          /** Optional sanitized vector projection. */
          readonly vector?: {
            /** Source vector viewport as x, y, width and height. */
            readonly viewBox: readonly [number, number, number, number];
            /** Sanitized SVG path data in paint order. */
            readonly paths: readonly string[];
            /** Foreground paint color. */
            readonly foreground: string;
            /** Background paint color. */
            readonly background: string;
          };
        }
      >
    >;
  };
}

/** One SVG page produced without consulting mutable editor state. */
export interface BrowserPrintSvgPage {
  /** Stable identity shared by preview and browser print. */
  readonly id: string;
  /** Page width in print-profile units. */
  readonly width: number;
  /** Page height in print-profile units. */
  readonly height: number;
  /** Complete standalone SVG markup. */
  readonly svg: string;
}

/** Browser print request options. */
export interface BrowserPrintOptions {
  /** Cancels the pending browser print session and removes its iframe. */
  readonly signal?: AbortSignal;
}

/** Why the isolated print document was removed. */
export type BrowserPrintCleanupReason = 'afterprint' | 'timeout';

/** Successful browser print invocation metadata. */
export interface BrowserPrintResult {
  /** Stable page identities sent to the browser. */
  readonly pageIds: readonly string[];
  /** Number of SVG pages sent to the browser. */
  readonly pageCount: number;
  /** Lifecycle event which completed cleanup. */
  readonly cleanupReason: BrowserPrintCleanupReason;
}

/** Stable browser-print failure code. */
export type BrowserPrintErrorCode = 'PRINT_BLOCKED' | 'RENDER_ABORTED';

/** Error raised when browser printing cannot start or is cancelled. */
export class BrowserPrintError extends Error {
  /** Stable machine-readable failure code. */
  readonly code: BrowserPrintErrorCode;

  /** Creates a stable browser-print failure. */
  constructor(code: BrowserPrintErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'BrowserPrintError';
    this.code = code;
  }
}

/** Construction options for the isolated browser adapter. */
export interface IsolatedBrowserPrintAdapterOptions {
  /** DOM document which owns the isolated iframe. */
  readonly document?: Document;
  /** Bounded afterprint fallback in milliseconds. */
  readonly timeoutMs?: number;
  /** Optional host print capability, primarily useful for embedded browsers. */
  readonly print?: (target: Window) => void;
}

interface PrintSession {
  readonly frame: HTMLIFrameElement;
  readonly target: Window;
  readonly afterPrint: () => void;
  readonly abort?: () => void;
  readonly signal?: AbortSignal;
  readonly timer: ReturnType<typeof setTimeout>;
  readonly resolve: (result: BrowserPrintResult) => void;
  readonly reject: (error: BrowserPrintError) => void;
  readonly pageIds: readonly string[];
}

function escapeMarkup(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function number(value: number): string {
  if (!Number.isFinite(value)) throw new RangeError('SVG geometry must be finite');
  return String(value);
}

function rectAttributes(rect: DisplayRect): string {
  return `x="${number(rect.x)}" y="${number(rect.y)}" width="${number(rect.width)}" height="${number(rect.height)}"`;
}

function safeLinkHref(href: string): string | undefined {
  if (href.startsWith('#') || /^(?:https?:|mailto:)/iu.test(href)) return href;
  return undefined;
}

function base64(bytes: readonly number[]): string {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    binary += String.fromCharCode(...bytes.slice(offset, offset + 32_768));
  }
  return btoa(binary);
}

function preserveAspectRatio(fit: 'contain' | 'cover' | 'fill'): string {
  return fit === 'fill' ? 'none' : fit === 'cover' ? 'xMidYMid slice' : 'xMidYMid meet';
}

function serializeCommands(
  commands: readonly PrintDisplayCommand[],
  path: string,
  resources: GeneratedDocumentForBrowserPrint['resources'],
): string {
  return commands
    .map((command, index) => {
      const commandPath = `${path}-${index}`;
      switch (command.kind) {
        case 'fill-rect':
          return `<rect ${rectAttributes(command.rect)} fill="${escapeMarkup(command.color)}"/>`;
        case 'stroke-rect':
          return `<rect ${rectAttributes(command.rect)} fill="none" stroke="${escapeMarkup(command.color)}" stroke-width="${number(command.width)}"/>`;
        case 'line':
          return `<line x1="${number(command.x1)}" y1="${number(command.y1)}" x2="${number(command.x2)}" y2="${number(command.y2)}" stroke="${escapeMarkup(command.color)}" stroke-width="${number(command.width)}"/>`;
        case 'text': {
          const anchor =
            command.horizontalAlign === 'center'
              ? 'middle'
              : command.horizontalAlign === 'right'
                ? 'end'
                : 'start';
          return `<text x="${number(command.x)}" y="${number(command.y)}" fill="${escapeMarkup(command.color)}" font-family="${escapeMarkup(command.fontFamily)}" font-size="${number(command.fontSize)}" text-anchor="${anchor}" data-max-width="${number(command.maxWidth)}">${escapeMarkup(command.text)}</text>`;
        }
        case 'path':
          return `<path d="${escapeMarkup(command.data)}"${command.fill === undefined ? ' fill="none"' : ` fill="${escapeMarkup(command.fill)}"`}${command.stroke === undefined ? '' : ` stroke="${escapeMarkup(command.stroke)}"`}${command.width === undefined ? '' : ` stroke-width="${number(command.width)}"`}/>`;
        case 'clip': {
          const clipId = `tego-clip-${commandPath}`;
          return `<defs><clipPath id="${clipId}"><rect ${rectAttributes(command.rect)}/></clipPath></defs><g clip-path="url(#${clipId})">${serializeCommands(command.commands, commandPath, resources)}</g>`;
        }
        case 'group':
          return `<g transform="rotate(${number(command.rotation)} ${number(command.origin.x)} ${number(command.origin.y)})">${serializeCommands(command.commands, commandPath, resources)}</g>`;
        case 'link': {
          const href = safeLinkHref(command.href);
          const body = `<rect ${rectAttributes(command.rect)} fill="transparent" aria-label="${escapeMarkup(command.label)}"/>`;
          return href === undefined
            ? body
            : `<a href="${escapeMarkup(href)}" aria-label="${escapeMarkup(command.label)}">${body}</a>`;
        }
        case 'image': {
          const resource = resources?.byReference[command.resourceId];
          if (resource?.vector !== undefined) {
            const vector = resource.vector;
            const background =
              vector.background === 'transparent'
                ? ''
                : `<rect x="${number(vector.viewBox[0])}" y="${number(vector.viewBox[1])}" width="${number(vector.viewBox[2])}" height="${number(vector.viewBox[3])}" fill="${escapeMarkup(vector.background)}"/>`;
            const paths = vector.paths
              .map(
                (data) =>
                  `<path d="${escapeMarkup(data)}" fill="${escapeMarkup(vector.foreground)}"/>`,
              )
              .join('');
            return `<svg ${rectAttributes(command.rect)} viewBox="${vector.viewBox.map(number).join(' ')}" preserveAspectRatio="${preserveAspectRatio(command.fit)}" data-resource-id="${escapeMarkup(command.resourceId)}">${background}${paths}</svg>`;
          }
          if (resource !== undefined && resource.type === 'image') {
            const href = `data:${resource.mimeType};base64,${base64(resource.bytes)}`;
            return `<image ${rectAttributes(command.rect)} href="${escapeMarkup(href)}" preserveAspectRatio="${preserveAspectRatio(command.fit)}" data-resource-id="${escapeMarkup(command.resourceId)}"/>`;
          }
          return `<g data-resource-id="${escapeMarkup(command.resourceId)}" data-fit="${command.fit}"><rect ${rectAttributes(command.rect)} fill="none"/></g>`;
        }
      }
    })
    .join('');
}

/** Serializes the exact immutable pages consumed by preview and browser print. */
export function serializeGeneratedDocumentSvgPages(
  document: GeneratedDocumentForBrowserPrint,
): readonly BrowserPrintSvgPage[] {
  return Object.freeze(
    document.print.displayList.pages.map((page) => {
      const semantic = document.print.pages[page.index];
      if (
        semantic === undefined ||
        semantic.width !== page.width ||
        semantic.height !== page.height
      ) {
        throw new TypeError('Generated print pages and display list are inconsistent');
      }
      const id = semantic.id;
      return Object.freeze({
        id,
        width: page.width,
        height: page.height,
        svg:
          `<svg xmlns="http://www.w3.org/2000/svg" data-page-id="${escapeMarkup(id)}" ` +
          `width="${number(page.width)}" height="${number(page.height)}" ` +
          `viewBox="0 0 ${number(page.width)} ${number(page.height)}" role="img">` +
          `${serializeCommands(page.commands, `page-${page.index}`, document.resources)}</svg>`,
      });
    }),
  );
}

function printHtml(pages: readonly BrowserPrintSvgPage[]): string {
  const pageMarkup = pages
    .map(
      (page) =>
        `<section class="tego-print-page" data-page-id="${escapeMarkup(page.id)}">${page.svg}</section>`,
    )
    .join('');
  return (
    '<!doctype html><html><head><meta charset="utf-8">' +
    '<style>@page{margin:0}html,body{margin:0;padding:0}.tego-print-page{break-after:page;page-break-after:always;overflow:hidden}.tego-print-page:last-child{break-after:auto;page-break-after:auto}svg{display:block}</style>' +
    `</head><body>${pageMarkup}</body></html>`
  );
}

function abortedError(message = 'Browser print was cancelled'): BrowserPrintError {
  return new BrowserPrintError('RENDER_ABORTED', message);
}

/**
 * Prints immutable SVG pages in an isolated same-origin iframe.
 *
 * The adapter never clones the editor DOM and owns every iframe it creates.
 */
export class IsolatedBrowserPrintAdapter {
  /** Owner document used only to mount isolated frames. */
  readonly #document: Document;
  /** Bounded cleanup fallback. */
  readonly #timeoutMs: number;
  /** Optional host print invocation. */
  readonly #print?: (target: Window) => void;
  /** Active sessions owned by this adapter. */
  readonly #sessions = new Set<PrintSession>();
  /** Whether disposal has permanently closed this adapter. */
  #disposed = false;

  /** Creates an iframe-owning browser print adapter. */
  constructor(options: IsolatedBrowserPrintAdapterOptions = {}) {
    const ownerDocument = options.document ?? globalThis.document;
    if (ownerDocument === undefined) {
      throw new TypeError('Browser printing requires a DOM Document');
    }
    const timeoutMs = options.timeoutMs ?? 30_000;
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
      throw new RangeError('Browser print timeout must be a positive finite number');
    }
    this.#document = ownerDocument;
    this.#timeoutMs = timeoutMs;
    this.#print = options.print;
  }

  /** Starts one isolated browser-print lifecycle. */
  print(
    document: GeneratedDocumentForBrowserPrint,
    options: BrowserPrintOptions = {},
  ): Promise<BrowserPrintResult> {
    if (this.#disposed) return Promise.reject(abortedError('Browser print adapter is disposed'));
    if (options.signal?.aborted === true) return Promise.reject(abortedError());

    const pages = serializeGeneratedDocumentSvgPages(document);
    const frame = this.#document.createElement('iframe');
    frame.hidden = true;
    frame.setAttribute('aria-hidden', 'true');
    frame.setAttribute('data-tego-browser-print', '');
    frame.src = 'about:blank';
    this.#document.body.append(frame);

    const target = frame.contentWindow;
    const targetDocument = frame.contentDocument;
    if (target === null || targetDocument === null) {
      frame.remove();
      return Promise.reject(
        new BrowserPrintError('PRINT_BLOCKED', 'Browser print iframe is unavailable'),
      );
    }

    try {
      targetDocument.open();
      targetDocument.write(printHtml(pages));
      targetDocument.close();
    } catch (cause) {
      frame.remove();
      return Promise.reject(
        new BrowserPrintError('PRINT_BLOCKED', 'Browser print document was blocked', { cause }),
      );
    }

    return new Promise<BrowserPrintResult>((resolve, reject) => {
      const finish = (cleanupReason: BrowserPrintCleanupReason): void => {
        if (!this.#sessions.has(session)) return;
        this.#cleanup(session);
        resolve(
          Object.freeze({
            pageIds: session.pageIds,
            pageCount: session.pageIds.length,
            cleanupReason,
          }),
        );
      };
      const fail = (error: BrowserPrintError): void => {
        if (!this.#sessions.has(session)) return;
        this.#cleanup(session);
        reject(error);
      };
      const afterPrint = (): void => finish('afterprint');
      const abort = options.signal === undefined ? undefined : (): void => fail(abortedError());
      const timer = setTimeout(() => finish('timeout'), this.#timeoutMs);
      const session: PrintSession = {
        frame,
        target,
        afterPrint,
        ...(abort === undefined ? {} : { abort }),
        ...(options.signal === undefined ? {} : { signal: options.signal }),
        timer,
        resolve,
        reject,
        pageIds: Object.freeze(pages.map((page) => page.id)),
      };
      this.#sessions.add(session);
      target.addEventListener('afterprint', afterPrint, { once: true });
      if (abort !== undefined) options.signal?.addEventListener('abort', abort, { once: true });

      const fontsReady = targetDocument.fonts?.ready ?? Promise.resolve();
      void fontsReady.then(
        () => {
          if (!this.#sessions.has(session)) return;
          try {
            const print = this.#print ?? ((printTarget: Window) => printTarget.print());
            print(target);
          } catch (cause) {
            fail(
              new BrowserPrintError('PRINT_BLOCKED', 'Browser print was blocked by the host', {
                cause,
              }),
            );
          }
        },
        (cause: unknown) => {
          fail(
            new BrowserPrintError('PRINT_BLOCKED', 'Browser print resources could not load', {
              cause,
            }),
          );
        },
      );
    });
  }

  /** Cancels active sessions and permanently releases this adapter. */
  dispose(): void {
    if (this.#disposed) return;
    this.#disposed = true;
    for (const session of this.#sessions) {
      this.#cleanup(session);
      session.reject(abortedError('Browser print adapter was disposed'));
    }
  }

  /** Removes listeners, timers, and the iframe for one owned session. */
  #cleanup(session: PrintSession): void {
    if (!this.#sessions.delete(session)) return;
    clearTimeout(session.timer);
    session.target.removeEventListener('afterprint', session.afterPrint);
    if (session.abort !== undefined) session.signal?.removeEventListener('abort', session.abort);
    session.frame.remove();
  }
}
