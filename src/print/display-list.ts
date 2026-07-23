import type { Diagnostic } from '../document';
import type { FontMetrics } from '../presentation';
import type { CellPresentation, ResolvedStyle } from '../presentation';

/** Device-independent rectangle in print-profile units. */
export interface DisplayRect {
  /** Left position in device-independent units. */
  readonly x: number;
  /** Top position in device-independent units. */
  readonly y: number;
  /** Non-negative width in device-independent units. */
  readonly width: number;
  /** Non-negative height in device-independent units. */
  readonly height: number;
}

/** Immutable rendering operation understood by output adapters. */
export type PrintDisplayCommand =
  | PrintFillRectCommand
  | PrintStrokeRectCommand
  | PrintTextCommand
  | PrintLineCommand
  | PrintImageCommand
  | PrintPathCommand
  | PrintClipCommand
  | PrintLinkCommand;

/** Filled rectangular display-list operation. */
export interface PrintFillRectCommand {
  /** Stable operation discriminator. */
  readonly kind: 'fill-rect';
  /** Rectangle to fill. */
  readonly rect: DisplayRect;
  /** Fill color. */
  readonly color: string;
}

/** Stroked rectangular display-list operation. */
export interface PrintStrokeRectCommand {
  /** Stable operation discriminator. */
  readonly kind: 'stroke-rect';
  /** Rectangle to stroke. */
  readonly rect: DisplayRect;
  /** Stroke color. */
  readonly color: string;
  /** Stroke width in device-independent units. */
  readonly width: number;
}

/** Plain-text display-list operation. */
export interface PrintTextCommand {
  /** Stable operation discriminator. */
  readonly kind: 'text';
  /** Text-node content; never interpreted as markup. */
  readonly text: string;
  /** Horizontal anchor in device-independent units. */
  readonly x: number;
  /** Vertical anchor in device-independent units. */
  readonly y: number;
  /** Maximum layout width. */
  readonly maxWidth: number;
  /** Resolved output font family. */
  readonly fontFamily: string;
  /** Resolved output font size. */
  readonly fontSize: number;
  /** Text color. */
  readonly color: string;
  /** Horizontal alignment around the anchor. */
  readonly horizontalAlign: ResolvedStyle['horizontalAlign'];
}

/** Device-independent line segment. */
export interface PrintLineCommand {
  /** Stable operation discriminator. */
  readonly kind: 'line';
  /** Starting horizontal coordinate. */
  readonly x1: number;
  /** Starting vertical coordinate. */
  readonly y1: number;
  /** Ending horizontal coordinate. */
  readonly x2: number;
  /** Ending vertical coordinate. */
  readonly y2: number;
  /** Stroke color. */
  readonly color: string;
  /** Stroke width. */
  readonly width: number;
}

/** Positioned immutable image resource reference. */
export interface PrintImageCommand {
  /** Stable operation discriminator. */
  readonly kind: 'image';
  /** Stable document resource identifier. */
  readonly resourceId: string;
  /** Image destination rectangle. */
  readonly rect: DisplayRect;
  /** Deterministic image fitting policy. */
  readonly fit: 'contain' | 'cover' | 'fill';
}

/** Device-independent vector path operation. */
export interface PrintPathCommand {
  /** Stable operation discriminator. */
  readonly kind: 'path';
  /** SVG-compatible path data. */
  readonly data: string;
  /** Optional fill color. */
  readonly fill?: string;
  /** Optional stroke color. */
  readonly stroke?: string;
  /** Optional stroke width. */
  readonly width?: number;
}

/** Nested clipping operation with structurally balanced commands. */
export interface PrintClipCommand {
  /** Stable operation discriminator. */
  readonly kind: 'clip';
  /** Clipping rectangle. */
  readonly rect: DisplayRect;
  /** Commands evaluated inside the clip. */
  readonly commands: readonly PrintDisplayCommand[];
}

/** Accessible link region in generated output. */
export interface PrintLinkCommand {
  /** Stable operation discriminator. */
  readonly kind: 'link';
  /** Link hit rectangle. */
  readonly rect: DisplayRect;
  /** Explicit destination URI. */
  readonly href: string;
  /** Accessible link label. */
  readonly label: string;
}

/** One immutable page in a print display list. */
export interface PrintDisplayPage {
  /** Zero-based stable page index. */
  readonly index: number;
  /** Page width in device-independent units. */
  readonly width: number;
  /** Page height in device-independent units. */
  readonly height: number;
  /** Ordered immutable rendering commands. */
  readonly commands: readonly PrintDisplayCommand[];
}

/** Deterministic, renderer-neutral print output. */
export interface PrintDisplayList {
  /** Ordered immutable output pages. */
  readonly pages: readonly PrintDisplayPage[];
  /** Recoverable layout and resource diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
}

/** Positioned presentation used to construct a display list. */
export interface PrintDisplayCell {
  /** Positioned cell rectangle. */
  readonly rect: DisplayRect;
  /** Shared cell semantics consumed without recalculation. */
  readonly presentation: CellPresentation;
}

/** Page input for display-list generation. */
export interface PrintDisplayPageInput {
  /** Page width in device-independent units. */
  readonly width: number;
  /** Page height in device-independent units. */
  readonly height: number;
  /** Positioned shared cell presentations. */
  readonly cells: readonly PrintDisplayCell[];
}

/** Explicit inputs for deterministic display-list generation. */
export interface PrintDisplayListInput {
  /** Ordered page inputs independent of the screen viewport. */
  readonly pages: readonly PrintDisplayPageInput[];
  /** Explicit deterministic font measurements. */
  readonly fontMetrics: FontMetrics;
}

function finiteNonNegative(value: number, name: string): void {
  if (!Number.isFinite(value) || value < 0) throw new RangeError(`${name} must be non-negative`);
}

function freezeCommand(command: PrintDisplayCommand): PrintDisplayCommand {
  if (command.kind === 'clip') {
    return Object.freeze({
      ...command,
      rect: Object.freeze({ ...command.rect }),
      commands: Object.freeze(command.commands.map(freezeCommand)),
    });
  }
  return Object.freeze({
    ...command,
    ...('rect' in command ? { rect: Object.freeze({ ...command.rect }) } : {}),
  });
}

const DRAW_COMMAND_KINDS = new Set([
  'fill-rect',
  'stroke-rect',
  'text',
  'line',
  'image',
  'path',
  'clip',
  'link',
]);

/** Diagnoses unsupported commands before an output adapter consumes a display list. */
export function validatePrintDisplayCommands(commands: readonly unknown[]): readonly Diagnostic[] {
  const diagnostics: Diagnostic[] = [];
  const visit = (command: unknown, index: number): void => {
    if (command === null || typeof command !== 'object' || Array.isArray(command)) {
      diagnostics.push({
        code: 'DRAW_COMMAND_UNSUPPORTED',
        severity: 'error',
        domain: 'output',
        stage: 'validate',
        message: `Draw command ${index} is not an object`,
      });
      return;
    }
    const candidate = command as Readonly<Record<string, unknown>>;
    if (typeof candidate.kind !== 'string' || !DRAW_COMMAND_KINDS.has(candidate.kind)) {
      diagnostics.push({
        code: 'DRAW_COMMAND_UNSUPPORTED',
        severity: 'error',
        domain: 'output',
        stage: 'validate',
        message: `Draw command ${index} has unsupported kind ${String(candidate.kind)}`,
        details: { index, kind: String(candidate.kind) },
      });
      return;
    }
    if (candidate.kind === 'clip' && Array.isArray(candidate.commands)) {
      candidate.commands.forEach(visit);
    }
  };
  commands.forEach(visit);
  return Object.freeze(diagnostics.map((diagnostic) => Object.freeze(diagnostic)));
}

function wrapText(
  text: string,
  maximumWidth: number,
  fontFamily: string,
  fontSize: number,
  metrics: FontMetrics,
): readonly string[] {
  if (text === '') return [];
  const lines: string[] = [];
  for (const sourceLine of text.split('\n')) {
    let line = '';
    for (const character of sourceLine) {
      const candidate = `${line}${character}`;
      if (line !== '' && metrics.measure(candidate, fontFamily, fontSize) > maximumWidth) {
        lines.push(line);
        line = character;
      } else {
        line = candidate;
      }
    }
    lines.push(line);
  }
  return lines;
}

/** Builds output commands using presentation text and injected font metrics only. */
export function createPrintDisplayList(input: PrintDisplayListInput): PrintDisplayList {
  const diagnostics: Diagnostic[] = [];
  const missingFonts = new Set<string>();
  const pages = input.pages.map((page, index): PrintDisplayPage => {
    finiteNonNegative(page.width, 'page width');
    finiteNonNegative(page.height, 'page height');
    const commands: PrintDisplayCommand[] = [];
    for (const { rect, presentation } of page.cells) {
      finiteNonNegative(rect.width, 'cell width');
      finiteNonNegative(rect.height, 'cell height');
      if (presentation.visibility.hidden) continue;
      commands.push({
        kind: 'fill-rect',
        rect: {
          x: rect.x + 1,
          y: rect.y + 1,
          width: Math.max(0, rect.width - 2),
          height: Math.max(0, rect.height - 2),
        },
        color: presentation.style.backgroundColor,
      });
      commands.push({ kind: 'stroke-rect', rect, color: '#d0d0d0', width: 1 });
      if (!presentation.visibility.printable) continue;
      if (presentation.accessibility.invalid) {
        diagnostics.push({
          code: 'PRESENTATION_CELL_INVALID',
          severity: 'error',
          domain: 'layout',
          stage: 'render',
          message: presentation.accessibility.description ?? 'Cell presentation is invalid',
          location: { cell: presentation.address },
        });
      }
      for (const annotation of presentation.annotations) {
        diagnostics.push({
          code: 'PRESENTATION_ANNOTATION',
          severity: 'info',
          domain: 'output',
          stage: 'render',
          message: annotation.text,
          location: { cell: presentation.address },
          details: { kind: annotation.kind },
        });
      }
      if (presentation.diagnostics !== undefined) diagnostics.push(...presentation.diagnostics);
      if (presentation.formattedText === '') continue;
      const font = input.fontMetrics.resolve(presentation.style.fontFamily);
      if (font.missing && !missingFonts.has(presentation.style.fontFamily)) {
        missingFonts.add(presentation.style.fontFamily);
        diagnostics.push({
          code: 'FONT_METRICS_UNAVAILABLE',
          severity: 'error',
          domain: 'layout',
          stage: 'resolve',
          message: `Font ${presentation.style.fontFamily} is unavailable`,
          details: {
            requestedFont: presentation.style.fontFamily,
            fallbackFont: font.fontFamily,
          },
        });
      }
      if (font.missing) continue;
      const padding = Math.min(5, rect.width / 2);
      const x =
        presentation.style.horizontalAlign === 'right'
          ? rect.x + rect.width - padding
          : presentation.style.horizontalAlign === 'center'
            ? rect.x + rect.width / 2
            : rect.x + padding;
      const maximumWidth = Math.max(0, rect.width - padding * 2);
      const lines = presentation.style.wrap
        ? wrapText(
            presentation.formattedText,
            maximumWidth,
            font.fontFamily,
            presentation.style.fontSize,
            input.fontMetrics,
          )
        : presentation.formattedText.split('\n');
      const lineHeight = font.lineHeight * (presentation.style.fontSize / 10);
      const firstY = rect.y + rect.height / 2 - ((lines.length - 1) * lineHeight) / 2;
      for (const [lineIndex, text] of lines.entries()) {
        commands.push({
          kind: 'text',
          text,
          x,
          y: firstY + lineIndex * lineHeight,
          maxWidth: maximumWidth,
          fontFamily: font.fontFamily,
          fontSize: presentation.style.fontSize,
          color: presentation.style.color,
          horizontalAlign: presentation.style.horizontalAlign,
        });
      }
    }
    return Object.freeze({
      index,
      width: page.width,
      height: page.height,
      commands: Object.freeze(commands.map(freezeCommand)),
    });
  });
  return Object.freeze({
    pages: Object.freeze(pages),
    diagnostics: Object.freeze(diagnostics.map((diagnostic) => Object.freeze(diagnostic))),
  });
}
