import type { ScreenObjectProjection } from '../../objects';
import type { PrintDisplayCommand } from '../../print';
import type { DrawContext } from './draw-context';

/** Paints the already-sorted screen object layer without changing persistent object state. */
export function paintObjects(
  draw: DrawContext,
  objects: readonly ScreenObjectProjection[],
  selectedObjectId?: string,
): void {
  for (const projection of objects) {
    paintCommands(draw.context, projection.commands, projection.imageResources);
    if (projection.object.id === selectedObjectId) {
      const { bounds } = projection;
      draw.context.save();
      try {
        draw.context.strokeStyle = '#2563eb';
        draw.context.lineWidth = 1;
        draw.context.setLineDash([4, 2]);
        draw.context.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
      } finally {
        draw.context.restore();
      }
    }
  }
}

function paintCommands(
  context: CanvasRenderingContext2D,
  commands: readonly PrintDisplayCommand[],
  imageResources?: ScreenObjectProjection['imageResources'],
): void {
  for (const command of commands) {
    if (command.kind === 'fill-rect') {
      context.fillStyle = command.color;
      context.fillRect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
      continue;
    }
    if (command.kind === 'stroke-rect') {
      context.strokeStyle = command.color;
      context.lineWidth = command.width;
      context.setLineDash([]);
      context.strokeRect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
      continue;
    }
    if (command.kind === 'line') {
      context.beginPath();
      context.moveTo(command.x1, command.y1);
      context.lineTo(command.x2, command.y2);
      context.strokeStyle = command.color;
      context.lineWidth = command.width;
      context.setLineDash([]);
      context.stroke();
      continue;
    }
    if (command.kind === 'text') {
      context.fillStyle = command.color;
      context.font = `${command.fontSize}px ${command.fontFamily}`;
      context.textAlign = command.horizontalAlign;
      context.textBaseline = 'alphabetic';
      context.fillText(command.text, command.x, command.y, command.maxWidth);
      continue;
    }
    if (command.kind === 'image') {
      const image = imageResources?.[command.resourceId];
      if (image === undefined) {
        paintImagePlaceholder(context, command.rect);
      } else {
        const placement = imagePlacement(image.width, image.height, command.rect, command.fit);
        context.drawImage(image.source as CanvasImageSource, ...placement);
      }
      continue;
    }
    if (command.kind === 'path') {
      paintPath(context, command);
      continue;
    }
    if (command.kind === 'clip') {
      context.save();
      try {
        context.beginPath();
        context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
        context.clip();
        paintCommands(context, command.commands, imageResources);
      } finally {
        context.restore();
      }
      continue;
    }
    if (command.kind === 'group') {
      context.save();
      try {
        context.translate(command.origin.x, command.origin.y);
        context.rotate((command.rotation * Math.PI) / 180);
        context.translate(-command.origin.x, -command.origin.y);
        paintCommands(context, command.commands, imageResources);
      } finally {
        context.restore();
      }
    }
  }
}

function imagePlacement(
  sourceWidth: number,
  sourceHeight: number,
  target: PrintDisplayCommandRect,
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

function paintImagePlaceholder(
  context: CanvasRenderingContext2D,
  rect: PrintDisplayCommandRect,
): void {
  context.fillStyle = '#e5e7eb';
  context.fillRect(rect.x, rect.y, rect.width, rect.height);
  context.strokeStyle = '#9ca3af';
  context.lineWidth = 1;
  context.setLineDash([]);
  context.strokeRect(rect.x, rect.y, rect.width, rect.height);
}

type PrintDisplayCommandRect = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

function paintPath(
  context: CanvasRenderingContext2D,
  command: Extract<PrintDisplayCommand, { readonly kind: 'path' }>,
): void {
  if (typeof Path2D === 'undefined') return;
  const path = new Path2D(command.data);
  if (command.fill !== undefined) {
    context.fillStyle = command.fill;
    context.fill(path);
  }
  if (command.stroke !== undefined) {
    context.strokeStyle = command.stroke;
    context.lineWidth = command.width ?? 1;
    context.setLineDash([]);
    context.stroke(path);
  }
}
