import type { DocumentSheetId, ResourceId, ResourceMetadata, SheetObject } from '../document';
import { archiveXml } from './archive';
import { InterchangeError, type ResolvedInterchangeLimits } from './contracts';
import { attributes, textContent } from './xml';

const EMU_PER_DIP = 9_525;

interface DrawingRelationship {
  readonly target: string;
  readonly external: boolean;
  readonly type: string;
}

export interface ParsedWorksheetDrawing {
  readonly objects: readonly SheetObject[];
  readonly resources: readonly ResourceMetadata[];
  readonly unsupported: readonly string[];
}

export interface XlsxDrawingResourcePool {
  readonly byMediaPath: Map<string, ResourceMetadata>;
  readonly contentByDigest: Map<
    string,
    {
      readonly bytes: Uint8Array;
      readonly resource: ResourceMetadata;
    }[]
  >;
  collisionComparisons: number;
  objectCount: number;
  resourceCount: number;
  resourceBytes: number;
  materializedBytes: number;
  readonly maxObjects: number;
  readonly maxResources: number;
  readonly maxResourceBytes: number;
  readonly maxMaterializedBytes: number;
}

export function createXlsxDrawingResourcePool(
  limits: Readonly<{
    readonly maxObjects?: number;
    readonly maxResources?: number;
    readonly maxResourceBytes?: number;
    readonly maxMaterializedBytes?: number;
  }> = {},
): XlsxDrawingResourcePool {
  return {
    byMediaPath: new Map(),
    contentByDigest: new Map(),
    collisionComparisons: 0,
    objectCount: 0,
    resourceCount: 0,
    resourceBytes: 0,
    materializedBytes: 0,
    maxObjects: limits.maxObjects ?? Number.MAX_SAFE_INTEGER,
    maxResources: limits.maxResources ?? Number.MAX_SAFE_INTEGER,
    maxResourceBytes: limits.maxResourceBytes ?? Number.MAX_SAFE_INTEGER,
    maxMaterializedBytes: limits.maxMaterializedBytes ?? Number.MAX_SAFE_INTEGER,
  };
}

function relationships(xml: string): ReadonlyMap<string, DrawingRelationship> {
  const result = new Map<string, DrawingRelationship>();
  for (const match of xml.matchAll(/<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/gi)) {
    const value = attributes(match[1]!);
    if (value.Id === undefined || value.Target === undefined) continue;
    result.set(value.Id, {
      target: value.Target,
      external:
        value.TargetMode?.toLowerCase() === 'external' || /^[a-z][\w+.-]*:/i.test(value.Target),
      type: value.Type ?? '',
    });
  }
  return result;
}

function relatedPart(basePart: string, target: string, requiredPrefix: string): string | undefined {
  if (target.startsWith('/') || target.startsWith('\\')) return undefined;
  const segments = [
    ...basePart.split('/').slice(0, -1),
    ...target.replaceAll('\\', '/').split('/'),
  ];
  const normalized: string[] = [];
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (normalized.length === 0) return undefined;
      normalized.pop();
    } else {
      normalized.push(segment);
    }
  }
  const result = normalized.join('/');
  return result.startsWith(requiredPrefix) ? result : undefined;
}

function relationshipsPart(part: string): string {
  const segments = part.split('/');
  const name = segments.pop()!;
  return `${segments.join('/')}/_rels/${name}.rels`;
}

function dip(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / EMU_PER_DIP : undefined;
}

function integerElement(body: string, name: 'col' | 'row'): number | undefined {
  const value = textContent(
    new RegExp(`<(?:[\\w.-]+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${name}>`, 'i').exec(
      body,
    )?.[1] ?? '',
  );
  const numeric = Number(value);
  return Number.isSafeInteger(numeric) && numeric >= 0 ? numeric : undefined;
}

function offsetElement(body: string, name: 'colOff' | 'rowOff'): number | undefined {
  const value = textContent(
    new RegExp(`<(?:[\\w.-]+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${name}>`, 'i').exec(
      body,
    )?.[1] ?? '',
  );
  return dip(value);
}

function marker(
  body: string,
  name: 'from' | 'to',
):
  | {
      readonly row: number;
      readonly column: number;
      readonly offset: { readonly x: number; readonly y: number };
    }
  | undefined {
  const content = new RegExp(
    `<(?:[\\w.-]+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${name}>`,
    'i',
  ).exec(body)?.[1];
  if (content === undefined) return undefined;
  const column = integerElement(content, 'col');
  const row = integerElement(content, 'row');
  const x = offsetElement(content, 'colOff');
  const y = offsetElement(content, 'rowOff');
  return column === undefined || row === undefined || x === undefined || y === undefined
    ? undefined
    : { row, column, offset: { x, y } };
}

function extent(body: string): { readonly width: number; readonly height: number } | undefined {
  const value = attributes(/<(?:[\w.-]+:)?ext\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '');
  const width = dip(value.cx);
  const height = dip(value.cy);
  return width === undefined || height === undefined || width < 0 || height < 0
    ? undefined
    : { width, height };
}

function anchor(
  type: 'absoluteAnchor' | 'oneCellAnchor' | 'twoCellAnchor',
  anchorAttributes: Readonly<Record<string, string>>,
  body: string,
  sheetId: DocumentSheetId,
): SheetObject['anchor'] | undefined {
  if (type === 'absoluteAnchor') {
    const position = attributes(/<(?:[\w.-]+:)?pos\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '');
    const x = dip(position.x);
    const y = dip(position.y);
    const size = extent(body);
    return x === undefined || y === undefined || size === undefined
      ? undefined
      : { type: 'absolute', rect: { x, y, ...size } };
  }
  const from = marker(body, 'from');
  if (from === undefined) return undefined;
  if (type === 'oneCellAnchor') {
    const size = extent(body);
    return size === undefined
      ? undefined
      : {
          type: 'one-cell',
          cell: { sheetId, row: from.row, column: from.column },
          offset: from.offset,
          size,
        };
  }
  const to = marker(body, 'to');
  if (to === undefined) return undefined;
  return {
    type: 'two-cell',
    from: { sheetId, ...from },
    to: { sheetId, ...to },
  };
}

function base64(bytes: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function base64Length(byteLength: number): number {
  return 4 * Math.ceil(byteLength / 3);
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

async function contentDigest(
  mimeType: 'image/png' | 'image/jpeg',
  bytes: Uint8Array,
): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes);
  const hex = [...new Uint8Array(digest)]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('');
  return `${mimeType}:sha256:${hex}`;
}

function limitExceeded(message: string): never {
  throw new InterchangeError('ARCHIVE_LIMIT_EXCEEDED', message);
}

function imageMime(path: string, bytes: Uint8Array): 'image/png' | 'image/jpeg' | undefined {
  if (
    path.toLowerCase().endsWith('.png') &&
    bytes.length >= 8 &&
    [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)
  ) {
    return 'image/png';
  }
  if (
    /\.(?:jpe?g)$/i.test(path) &&
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  ) {
    return 'image/jpeg';
  }
  return undefined;
}

function rotation(body: string): number | undefined {
  const value = Number(
    attributes(/<(?:[\w.-]+:)?xfrm\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '').rot,
  );
  if (!Number.isFinite(value)) return undefined;
  const degrees = (((value / 60_000) % 360) + 360) % 360;
  return degrees === 0 ? undefined : degrees;
}

function nonVisual(body: string): {
  readonly name: string;
  readonly description?: string;
  readonly locked: boolean;
} {
  const value = attributes(/<(?:[\w.-]+:)?cNvPr\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '');
  return {
    name: value.name ?? 'Imported object',
    ...(value.descr === undefined ? {} : { description: value.descr }),
    locked: /<(?:[\w.-]+:)?(?:spLocks|picLocks)\b[^>]*(?:noMove|noResize)="1"/i.test(body),
  };
}

function shapeStyle(body: string): {
  readonly fill?: string;
  readonly stroke?: string;
  readonly strokeWidth?: number;
} {
  const properties =
    /<(?:[\w.-]+:)?spPr\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?spPr>/i.exec(body)?.[1] ?? '';
  const line = /<(?:[\w.-]+:)?ln\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?ln>/i.exec(properties);
  const fillArea = line === null ? properties : properties.slice(0, line.index);
  const fill = attributes(
    /<(?:[\w.-]+:)?srgbClr\b([^>]*?)(?:\/>|>)/i.exec(fillArea)?.[1] ?? '',
  ).val;
  const stroke =
    line === null
      ? undefined
      : attributes(/<(?:[\w.-]+:)?srgbClr\b([^>]*?)(?:\/>|>)/i.exec(line[2]!)?.[1] ?? '').val;
  const width = line === null ? undefined : dip(attributes(line[1]!).w);
  return {
    ...(fill === undefined ? {} : { fill: `#${fill.toLowerCase()}` }),
    ...(stroke === undefined ? {} : { stroke: `#${stroke.toLowerCase()}` }),
    ...(width === undefined ? {} : { strokeWidth: width }),
  };
}

function textBoxStyle(body: string): Extract<SheetObject, { kind: 'text-box' }>['style'] {
  const runProperties = attributes(/<(?:[\w.-]+:)?rPr\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '');
  const font = attributes(/<(?:[\w.-]+:)?latin\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '').typeface;
  const color = attributes(/<(?:[\w.-]+:)?srgbClr\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '').val;
  const alignment = attributes(/<(?:[\w.-]+:)?pPr\b([^>]*?)(?:\/>|>)/i.exec(body)?.[1] ?? '').algn;
  return {
    color: color === undefined ? '#000000' : `#${color.toLowerCase()}`,
    fontFamily: font ?? 'Arial',
    fontSize: Number(runProperties.sz ?? 1_200) / 100,
    ...(alignment === undefined
      ? {}
      : {
          horizontalAlign:
            alignment === 'ctr'
              ? ('center' as const)
              : alignment === 'r'
                ? ('right' as const)
                : ('left' as const),
        }),
  };
}

function anchorEntries(xml: string): readonly {
  readonly index: number;
  readonly type: 'absoluteAnchor' | 'oneCellAnchor' | 'twoCellAnchor';
  readonly attributes: Readonly<Record<string, string>>;
  readonly body: string;
}[] {
  return (['absoluteAnchor', 'oneCellAnchor', 'twoCellAnchor'] as const)
    .flatMap((type) =>
      [
        ...xml.matchAll(
          new RegExp(
            `<(?:[\\w.-]+:)?${type}\\b([^>]*)>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${type}>`,
            'gi',
          ),
        ),
      ].map((match) => ({
        index: match.index,
        type,
        attributes: attributes(match[1]!),
        body: match[2]!,
      })),
    )
    .sort((left, right) => left.index - right.index);
}

function reserveDrawingObjects(xml: string, resourcePool: XlsxDrawingResourcePool): void {
  const openingAnchor = /<(?:[\w.-]+:)?(?:absoluteAnchor|oneCellAnchor|twoCellAnchor)\b/gi;
  let count = 0;
  while (openingAnchor.exec(xml) !== null) {
    count += 1;
    if (
      !Number.isSafeInteger(resourcePool.objectCount + count) ||
      resourcePool.objectCount + count > resourcePool.maxObjects
    ) {
      limitExceeded('XLSX DrawingML object count exceeds its workbook budget');
    }
  }
  resourcePool.objectCount += count;
}

export async function parseWorksheetDrawing(
  entries: Readonly<Record<string, Uint8Array>>,
  worksheetPart: string,
  worksheetXml: string,
  sheetId: DocumentSheetId,
  limits: ResolvedInterchangeLimits,
  resourcePool: XlsxDrawingResourcePool = createXlsxDrawingResourcePool(),
): Promise<ParsedWorksheetDrawing> {
  const unsupported: string[] = [];
  const drawingReference = attributes(
    /<(?:[\w.-]+:)?drawing\b([^>]*?)(?:\/>|>)/i.exec(worksheetXml)?.[1] ?? '',
  )['r:id'];
  if (drawingReference === undefined) return { objects: [], resources: [], unsupported };
  const worksheetRelationshipsPath = relationshipsPart(worksheetPart);
  const worksheetRelationshipsBytes = entries[worksheetRelationshipsPath];
  if (worksheetRelationshipsBytes === undefined) {
    return { objects: [], resources: [], unsupported: ['xlsx:drawing-objects'] };
  }
  const worksheetRelationships = relationships(
    archiveXml(entries, worksheetRelationshipsPath, limits),
  );
  const drawingRelationship = worksheetRelationships.get(drawingReference);
  if (drawingRelationship === undefined || drawingRelationship.external) {
    return { objects: [], resources: [], unsupported: ['xlsx:drawing-external-relationship'] };
  }
  if (!drawingRelationship.type.endsWith('/drawing') && drawingRelationship.type !== 'drawing') {
    return { objects: [], resources: [], unsupported: ['xlsx:drawing-resource-unsafe'] };
  }
  const drawingPart = relatedPart(worksheetPart, drawingRelationship.target, 'xl/drawings/');
  if (drawingPart === undefined || entries[drawingPart] === undefined) {
    return { objects: [], resources: [], unsupported: ['xlsx:drawing-resource-unsafe'] };
  }
  const drawingXml = archiveXml(entries, drawingPart, limits);
  const drawingRelationshipsPath = relationshipsPart(drawingPart);
  const drawingRelationships =
    entries[drawingRelationshipsPath] === undefined
      ? new Map<string, DrawingRelationship>()
      : relationships(archiveXml(entries, drawingRelationshipsPath, limits));
  const objects: SheetObject[] = [];
  const resources: ResourceMetadata[] = [];
  reserveDrawingObjects(drawingXml, resourcePool);
  const entriesInDrawing = anchorEntries(drawingXml);
  for (const [zIndex, entry] of entriesInDrawing.entries()) {
    if (
      entry.type === 'twoCellAnchor' &&
      entry.attributes.editAs !== undefined &&
      entry.attributes.editAs !== 'twoCell'
    ) {
      unsupported.push('xlsx:drawing-editas-unsupported');
      continue;
    }
    const objectAnchor = anchor(entry.type, entry.attributes, entry.body, sheetId);
    if (objectAnchor === undefined) {
      unsupported.push('xlsx:drawing-object-invalid');
      continue;
    }
    const accessible = nonVisual(entry.body);
    const common = {
      id: `xlsx-object-${zIndex + 1}`,
      anchor: objectAnchor,
      zIndex,
      locked: accessible.locked,
      templateRepeat: 'shared' as const,
      accessibility: {
        name: accessible.name,
        ...(accessible.description === undefined ? {} : { description: accessible.description }),
      },
      ...(rotation(entry.body) === undefined ? {} : { rotation: rotation(entry.body) }),
    };
    if (/<(?:[\w.-]+:)?pic\b/i.test(entry.body)) {
      const relationshipId = attributes(
        /<(?:[\w.-]+:)?blip\b([^>]*?)(?:\/>|>)/i.exec(entry.body)?.[1] ?? '',
      )['r:embed'];
      const relationship =
        relationshipId === undefined ? undefined : drawingRelationships.get(relationshipId);
      if (relationship === undefined || relationship.external) {
        unsupported.push('xlsx:drawing-external-relationship');
        continue;
      }
      if (!relationship.type.endsWith('/image') && relationship.type !== 'image') {
        unsupported.push('xlsx:drawing-resource-unsupported');
        continue;
      }
      const mediaPath = relatedPart(drawingPart, relationship.target, 'xl/media/');
      if (mediaPath === undefined || entries[mediaPath] === undefined) {
        unsupported.push('xlsx:drawing-resource-unsafe');
        continue;
      }
      const bytes = entries[mediaPath]!;
      const mimeType = imageMime(mediaPath, bytes);
      if (mimeType === undefined) {
        unsupported.push('xlsx:drawing-resource-unsupported');
        continue;
      }
      let resource = resourcePool.byMediaPath.get(mediaPath);
      let digest: string | undefined;
      if (resource === undefined) {
        const encodedBytes =
          (`data:${mimeType};base64,`.length + base64Length(bytes.byteLength)) * 2;
        if (
          resourcePool.maxResources === 0 ||
          bytes.byteLength > resourcePool.maxResourceBytes ||
          encodedBytes > resourcePool.maxMaterializedBytes
        ) {
          limitExceeded('XLSX drawing resource exceeds its workbook budget');
        }
        digest = await contentDigest(mimeType, bytes);
        resource = resourcePool.contentByDigest.get(digest)?.find((candidate) => {
          resourcePool.collisionComparisons += 1;
          return sameBytes(candidate.bytes, bytes);
        })?.resource;
      }
      if (resource === undefined) {
        const prefix = `data:${mimeType};base64,`;
        const materializedBytes = (prefix.length + base64Length(bytes.byteLength)) * 2;
        if (resourcePool.resourceCount + 1 > resourcePool.maxResources) {
          limitExceeded('XLSX drawing resource count exceeds its workbook budget');
        }
        if (
          !Number.isSafeInteger(resourcePool.resourceBytes + bytes.byteLength) ||
          resourcePool.resourceBytes + bytes.byteLength > resourcePool.maxResourceBytes
        ) {
          limitExceeded('XLSX drawing resource bytes exceed their workbook budget');
        }
        if (
          !Number.isSafeInteger(resourcePool.materializedBytes + materializedBytes) ||
          resourcePool.materializedBytes + materializedBytes > resourcePool.maxMaterializedBytes
        ) {
          limitExceeded('XLSX image materialization exceeds its memory budget');
        }
        const url = `${prefix}${base64(bytes)}`;
        resource = {
          id: `xlsx-resource-${resourcePool.byMediaPath.size + 1}` as ResourceId,
          kind: 'image',
          mimeType,
          byteLength: bytes.byteLength,
          url,
        };
        const bucket = resourcePool.contentByDigest.get(digest!);
        if (bucket === undefined) {
          resourcePool.contentByDigest.set(digest!, [{ bytes, resource }]);
        } else {
          bucket.push({ bytes, resource });
        }
        resourcePool.resourceCount += 1;
        resourcePool.resourceBytes += bytes.byteLength;
        resourcePool.materializedBytes += materializedBytes;
        resources.push(resource);
      }
      resourcePool.byMediaPath.set(mediaPath, resource);
      const resourceId = resource.id;
      objects.push({ ...common, kind: 'image', resourceId, fit: 'fill' } as SheetObject);
      continue;
    }
    const preset = attributes(
      /<(?:[\w.-]+:)?prstGeom\b([^>]*?)(?:\/>|>)/i.exec(entry.body)?.[1] ?? '',
    ).prst;
    if (/<(?:[\w.-]+:)?cNvSpPr\b[^>]*\btxBox="1"/i.test(entry.body)) {
      const text = [...entry.body.matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/gi)]
        .map((match) => textContent(match[1]!))
        .join('');
      objects.push({
        ...common,
        kind: 'text-box',
        text,
        style: textBoxStyle(entry.body),
      } as SheetObject);
    } else if (preset === 'rect' || preset === 'ellipse' || preset === 'line') {
      objects.push({
        ...common,
        kind: 'shape',
        shape: preset === 'rect' ? 'rectangle' : preset,
        style: shapeStyle(entry.body),
      } as SheetObject);
    } else {
      unsupported.push('xlsx:drawing-unknown-shape');
    }
  }
  return { objects, resources, unsupported };
}
