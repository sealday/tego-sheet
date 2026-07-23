/** Reads the OpenType OS/2 embedding flags from an SFNT or TTC font. */
export function openTypeEmbeddingFlags(bytesInput: readonly number[]): number | undefined {
  const bytes = new Uint8Array(bytesInput);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.byteLength < 12) return undefined;
  const tag = (offset: number): string =>
    String.fromCharCode(bytes[offset]!, bytes[offset + 1]!, bytes[offset + 2]!, bytes[offset + 3]!);
  let sfntOffset = 0;
  if (tag(0) === 'ttcf') {
    if (bytes.byteLength < 16 || view.getUint32(8) === 0) return undefined;
    sfntOffset = view.getUint32(12);
  }
  if (sfntOffset > bytes.byteLength - 12) return undefined;
  const tableCount = view.getUint16(sfntOffset + 4);
  const directoryEnd = sfntOffset + 12 + tableCount * 16;
  if (directoryEnd > bytes.byteLength) return undefined;
  for (let index = 0; index < tableCount; index += 1) {
    const record = sfntOffset + 12 + index * 16;
    if (tag(record) !== 'OS/2') continue;
    const tableOffset = view.getUint32(record + 8);
    const tableLength = view.getUint32(record + 12);
    if (tableLength < 10 || tableOffset > bytes.byteLength - tableLength) return undefined;
    return view.getUint16(tableOffset + 8);
  }
  return undefined;
}

/** Whether the OpenType policy forbids embedding or outline conversion. */
export function forbiddenFontEmbedding(flags: number): boolean {
  return (flags & 0x0002) !== 0 || (flags & 0x0200) !== 0;
}

/** Whether the OpenType policy forbids producing a used-glyph subset. */
export function forbiddenFontSubsetting(flags: number): boolean {
  return (flags & 0x0100) !== 0;
}
