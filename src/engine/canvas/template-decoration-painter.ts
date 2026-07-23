import type { FrozenQuadrantKind } from '../geometry/frozen-pane-geometry';
import { overlayAnchors } from '../geometry/overlay-anchors';
import type { ViewportMetrics } from '../ports';
import type { DrawContext } from './draw-context';
import type { TemplateCanvasDecoration } from './canvas-engine';

const COLORS = Object.freeze({
  value: '#2563eb',
  repeat: '#7c3aed',
  print: '#059669',
  error: '#dc2626',
});

export function paintTemplateDecorations(
  draw: DrawContext,
  decorations: readonly TemplateCanvasDecoration[],
  viewport: ViewportMetrics,
  pane: FrozenQuadrantKind,
): void {
  for (const decoration of decorations) {
    const color = decoration.invalid ? COLORS.error : COLORS[decoration.kind];
    for (const rect of overlayAnchors(decoration.range, viewport)) {
      if (rect.pane !== pane) continue;
      draw.strokeRect(rect, color, decoration.kind === 'print' ? 3 : 2);
      draw.fillRect(
        {
          left: rect.left,
          top: rect.top,
          width: Math.min(rect.width, Math.max(24, decoration.label.length * 7 + 8)),
          height: Math.min(18, rect.height),
        },
        color,
      );
      draw.text(
        decoration.label,
        { x: rect.left + 4, y: rect.top + 9 },
        {
          align: 'left',
          baseline: 'middle',
          color: '#ffffff',
          font: '11px Arial',
        },
      );
    }
  }
}
