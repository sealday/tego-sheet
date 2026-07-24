import { strFromU8, unzipSync } from 'fflate';
import {
  inputBytes,
  InterchangeError,
  throwIfAborted,
  type InterchangeInput,
  type ResolvedInterchangeLimits,
} from './contracts';
import { assertSafeXml } from './xml';

const decoder = new TextDecoder();

function uint16(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

function uint32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset]! |
      (bytes[offset + 1]! << 8) |
      (bytes[offset + 2]! << 16) |
      (bytes[offset + 3]! << 24)) >>>
    0
  );
}

function endOfCentralDirectory(bytes: Uint8Array): number {
  const minimum = Math.max(0, bytes.length - 65_557);
  for (let offset = bytes.length - 22; offset >= minimum; offset -= 1) {
    if (uint32(bytes, offset) === 0x06054b50) return offset;
  }
  throw new InterchangeError('ARCHIVE_INVALID', 'ZIP central directory was not found');
}

function inspectCentralDirectory(bytes: Uint8Array, limits: ResolvedInterchangeLimits): void {
  const end = endOfCentralDirectory(bytes);
  const entryCount = uint16(bytes, end + 10);
  const centralOffset = uint32(bytes, end + 16);
  if (entryCount > limits.maxEntries) {
    throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'ZIP entry limit exceeded');
  }
  let offset = centralOffset;
  let uncompressedBytes = 0;
  for (let index = 0; index < entryCount; index += 1) {
    if (uint32(bytes, offset) !== 0x02014b50) {
      throw new InterchangeError('ARCHIVE_INVALID', 'ZIP central directory is malformed');
    }
    const flags = uint16(bytes, offset + 8);
    const compressedSize = uint32(bytes, offset + 20);
    const uncompressedSize = uint32(bytes, offset + 24);
    const nameLength = uint16(bytes, offset + 28);
    const extraLength = uint16(bytes, offset + 30);
    const commentLength = uint16(bytes, offset + 32);
    if ((flags & 1) !== 0) {
      throw new InterchangeError(
        'UNSUPPORTED_ARCHIVE_FEATURE',
        'Encrypted ZIP entries are disabled',
      );
    }
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) {
      throw new InterchangeError('UNSUPPORTED_ARCHIVE_FEATURE', 'ZIP64 entries are not supported');
    }
    uncompressedBytes += uncompressedSize;
    if (uncompressedBytes > limits.maxUncompressedBytes) {
      throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'ZIP uncompressed byte limit exceeded');
    }
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (
      name.startsWith('/') ||
      name.startsWith('\\') ||
      name.split(/[\\/]/).some((segment) => segment === '..')
    ) {
      throw new InterchangeError('ARCHIVE_INVALID', 'Unsafe ZIP entry path');
    }
    offset += 46 + nameLength + extraLength + commentLength;
    if (offset > bytes.length) {
      throw new InterchangeError('ARCHIVE_INVALID', 'ZIP central directory exceeds package bounds');
    }
  }
}

export async function readArchive(
  input: InterchangeInput,
  limits: ResolvedInterchangeLimits,
  signal?: AbortSignal,
): Promise<Readonly<Record<string, Uint8Array>>> {
  const bytes = await inputBytes(input, signal, limits.maxPackageBytes);
  inspectCentralDirectory(bytes, limits);
  throwIfAborted(signal);
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(bytes);
  } catch (cause) {
    throw new InterchangeError('ARCHIVE_INVALID', 'ZIP package could not be decoded', { cause });
  }
  throwIfAborted(signal);
  let actualBytes = 0;
  for (const value of Object.values(entries)) {
    actualBytes += value.byteLength;
    if (actualBytes > limits.maxUncompressedBytes) {
      throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', 'ZIP uncompressed byte limit exceeded');
    }
  }
  return entries;
}

export function archiveXml(
  entries: Readonly<Record<string, Uint8Array>>,
  name: string,
  limits: ResolvedInterchangeLimits,
): string {
  const bytes = entries[name];
  if (!bytes)
    throw new InterchangeError('MALFORMED_WORKBOOK', `Missing required package part: ${name}`);
  if (bytes.byteLength > limits.maxXmlBytes) {
    throw new InterchangeError('XML_LIMIT_EXCEEDED', `XML part exceeds byte limit: ${name}`);
  }
  const xml = strFromU8(bytes);
  assertSafeXml(xml);
  return xml;
}
