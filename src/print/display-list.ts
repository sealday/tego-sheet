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
export type PrintDisplayCommand = PrintFillRectCommand | PrintStrokeRectCommand | PrintTextCommand;

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
  return Object.freeze({
    ...command,
    ...(command.kind === 'text' ? {} : { rect: Object.freeze({ ...command.rect }) }),
  });
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
      if (presentation.visibility.hidden || !presentation.visibility.printable) continue;
      commands.push({
        kind: 'fill-rect',
        rect,
        color: presentation.style.backgroundColor,
      });
      commands.push({ kind: 'stroke-rect', rect, color: '#d0d0d0', width: 1 });
      if (presentation.formattedText === '') continue;
      const font = input.fontMetrics.resolve(presentation.style.fontFamily);
      if (font.missing && !missingFonts.has(presentation.style.fontFamily)) {
        missingFonts.add(presentation.style.fontFamily);
        diagnostics.push({
          code: 'PRESENTATION_FONT_MISSING',
          severity: 'warning',
          domain: 'layout',
          stage: 'resolve',
          message: `Font ${presentation.style.fontFamily} is unavailable; using ${font.fontFamily}`,
          details: {
            requestedFont: presentation.style.fontFamily,
            fallbackFont: font.fontFamily,
          },
        });
      }
      const padding = Math.min(5, rect.width / 2);
      const x =
        presentation.style.horizontalAlign === 'right'
          ? rect.x + rect.width - padding
          : presentation.style.horizontalAlign === 'center'
            ? rect.x + rect.width / 2
            : rect.x + padding;
      commands.push({
        kind: 'text',
        text: presentation.formattedText,
        x,
        y: rect.y + rect.height / 2,
        maxWidth: Math.max(0, rect.width - padding * 2),
        fontFamily: font.fontFamily,
        fontSize: presentation.style.fontSize,
        color: presentation.style.color,
        horizontalAlign: presentation.style.horizontalAlign,
      });
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
