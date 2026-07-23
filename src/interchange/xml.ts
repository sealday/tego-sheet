import { InterchangeError } from './contracts';

const ENTITY_PATTERN = /&(?:#(\d+)|#x([\da-f]+)|([a-zA-Z][\w.-]*));/g;

export function assertSafeXml(xml: string): void {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new InterchangeError(
      'XML_ENTITY_REJECTED',
      'XML entities and document types are disabled',
    );
  }
}

export function decodeXml(value: string): string {
  return value.replace(ENTITY_PATTERN, (entity, decimal, hexadecimal, named) => {
    if (decimal !== undefined) return String.fromCodePoint(Number.parseInt(decimal, 10));
    if (hexadecimal !== undefined) return String.fromCodePoint(Number.parseInt(hexadecimal, 16));
    switch (named) {
      case 'amp':
        return '&';
      case 'apos':
        return "'";
      case 'gt':
        return '>';
      case 'lt':
        return '<';
      case 'quot':
        return '"';
      default:
        throw new InterchangeError('MALFORMED_WORKBOOK', `Unsupported XML entity: ${entity}`);
    }
  });
}

export function attributes(source: string): Readonly<Record<string, string>> {
  const result: Record<string, string> = Object.create(null);
  const pattern = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (const match of source.matchAll(pattern)) {
    result[match[1]!] = decodeXml(match[2] ?? match[3] ?? '');
  }
  return result;
}

export function textContent(source: string): string {
  return decodeXml(source.replace(/<[^>]*>/g, ''));
}
