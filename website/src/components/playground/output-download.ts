export function downloadBlob(
  blob: Blob,
  filename: string,
  targetDocument: Document = document,
  url: Pick<typeof URL, 'createObjectURL' | 'revokeObjectURL'> = URL,
): void {
  const href = url.createObjectURL(blob);
  const anchor = targetDocument.createElement('a');
  anchor.href = href;
  anchor.download = filename;
  anchor.hidden = true;
  targetDocument.body.append(anchor);
  try {
    anchor.click();
  } finally {
    anchor.remove();
    queueMicrotask(() => url.revokeObjectURL(href));
  }
}
