import type { CellStyle, SheetData } from '../core';
import {
  createPrintLayout,
  renderPrintPage,
  type PaperOrientation,
  type PaperSizeName,
} from '../engine';
import type { PrintDisplayCommand, PrintDisplayList } from '../print';
import type { FontMetrics } from '../presentation';

export interface PrintWorkbookOptions {
  readonly orientation: PaperOrientation;
  readonly paper: PaperSizeName;
  readonly fontMetrics?: FontMetrics;
}

function paintDisplayCommand(
  context: CanvasRenderingContext2D,
  command: PrintDisplayCommand,
): void {
  if (command.kind === 'fill-rect') {
    context.fillStyle = command.color;
    context.fillRect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
    return;
  }
  if (command.kind === 'stroke-rect') {
    context.beginPath();
    context.strokeStyle = command.color;
    context.lineWidth = command.width;
    context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
    context.stroke();
    return;
  }
  if (command.kind === 'line') {
    context.beginPath();
    context.strokeStyle = command.color;
    context.lineWidth = command.width;
    context.moveTo(command.x1, command.y1);
    context.lineTo(command.x2, command.y2);
    context.stroke();
    return;
  }
  if (command.kind === 'path') {
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
    return;
  }
  if (command.kind === 'clip') {
    context.save();
    context.beginPath();
    context.rect(command.rect.x, command.rect.y, command.rect.width, command.rect.height);
    context.clip();
    for (const nested of command.commands) paintDisplayCommand(context, nested);
    context.restore();
    return;
  }
  if (command.kind === 'image' || command.kind === 'link') return;
  context.fillStyle = command.color;
  context.font = `${command.fontSize}px ${command.fontFamily}`;
  context.textAlign = command.horizontalAlign;
  context.textBaseline = 'middle';
  context.fillText(command.text, command.x, command.y, command.maxWidth);
}

interface HiddenPrintSibling {
  readonly element: Element;
  readonly wasHidden: boolean;
}

function hidePrintSiblings(): readonly HiddenPrintSibling[] {
  const hidden: HiddenPrintSibling[] = [];
  try {
    for (const element of document.body.children) {
      const sibling = { element, wasHidden: element.hasAttribute('hidden') };
      hidden.push(sibling);
      element.setAttribute('hidden', '');
    }
    return hidden;
  } catch (error) {
    const rollbackErrors = restorePrintSiblings(hidden);
    if (rollbackErrors.length === 0) throw error;
    throw new AggregateError(
      [error, ...rollbackErrors],
      'Print sibling isolation and rollback both failed',
      { cause: error },
    );
  }
}

function restorePrintSiblings(siblings: readonly HiddenPrintSibling[]): unknown[] {
  const errors: unknown[] = [];
  for (const sibling of siblings) {
    if (sibling.wasHidden) continue;
    try {
      sibling.element.removeAttribute('hidden');
    } catch (error) {
      errors.push(error);
    }
  }
  return errors;
}

function removePrintNodes(host: HTMLElement, style: HTMLStyleElement): unknown[] {
  const errors: unknown[] = [];
  for (const node of [host, style]) {
    try {
      node.remove();
    } catch (error) {
      errors.push(error);
    }
  }
  return errors;
}

function throwCleanupErrors(errors: readonly unknown[]): void {
  if (errors.length === 1) throw errors[0];
  if (errors.length > 1) throw new AggregateError(errors, 'Print page cleanup failed');
}

export function mountPrintPages(
  sheet: SheetData,
  options: PrintWorkbookOptions,
  defaultStyle?: CellStyle,
): () => void {
  const layout = createPrintLayout(sheet, {
    paperSize: options.paper,
    orientation: options.orientation,
    defaultStyle,
  });
  const host = document.createElement('div');
  host.className = 'tego-sheet tego-sheet__print-pages';
  host.setAttribute('data-tego-print-pages', '');
  for (const page of layout.pages) {
    const canvas = document.createElement('canvas');
    canvas.setAttribute('data-page', String(page.index + 1));
    renderPrintPage(layout, page.index, canvas);
    host.append(canvas);
  }
  const style = document.createElement('style');
  style.setAttribute('data-tego-print-style', '');
  style.textContent = `@page { size: ${options.paper} ${options.orientation}; }`;
  // printWorkbook mounts and cleans synchronously. Remembering the prior hidden
  // attribute also makes nested mounts safe when they are released in LIFO order.
  let hiddenSiblings: readonly HiddenPrintSibling[] = [];
  try {
    document.head.append(style);
    hiddenSiblings = hidePrintSiblings();
    document.body.append(host);
  } catch (error) {
    const cleanupErrors = [
      ...removePrintNodes(host, style),
      ...restorePrintSiblings(hiddenSiblings),
    ];
    if (cleanupErrors.length === 0) throw error;
    throw new AggregateError(
      [error, ...cleanupErrors],
      'Print page installation and rollback failed',
      { cause: error },
    );
  }
  return () => {
    throwCleanupErrors([...removePrintNodes(host, style), ...restorePrintSiblings(hiddenSiblings)]);
  };
}

/** Mounts production print pages by translating an immutable display list only. */
export function mountPrintDisplayPages(
  displayList: PrintDisplayList,
  options: PrintWorkbookOptions,
): () => void {
  const host = document.createElement('div');
  host.className = 'tego-sheet tego-sheet__print-pages';
  host.setAttribute('data-tego-print-pages', '');
  for (const page of displayList.pages) {
    const canvas = document.createElement('canvas');
    canvas.width = page.width;
    canvas.height = page.height;
    canvas.setAttribute('data-page', String(page.index + 1));
    const context = canvas.getContext('2d');
    if (context === null) throw new Error('Print display-list Canvas context is unavailable');
    for (const command of page.commands) paintDisplayCommand(context, command);
    host.append(canvas);
  }
  const style = document.createElement('style');
  style.setAttribute('data-tego-print-style', '');
  style.textContent = `@page { size: ${options.paper} ${options.orientation}; }`;
  let hiddenSiblings: readonly HiddenPrintSibling[] = [];
  try {
    document.head.append(style);
    hiddenSiblings = hidePrintSiblings();
    document.body.append(host);
  } catch (error) {
    const cleanupErrors = [
      ...removePrintNodes(host, style),
      ...restorePrintSiblings(hiddenSiblings),
    ];
    if (cleanupErrors.length === 0) throw error;
    throw new AggregateError(
      [error, ...cleanupErrors],
      'Print display-list installation and rollback both failed',
      { cause: error },
    );
  }
  return () => {
    throwCleanupErrors([...removePrintNodes(host, style), ...restorePrintSiblings(hiddenSiblings)]);
  };
}
