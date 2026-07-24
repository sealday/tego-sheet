import type { Sheet, SheetObject } from '../../document';
import type { GeneratedDocument, ResolvedResource } from '../../template';
import { outputError } from '../output-error';

const EMU_PER_DIP = 9_525;

export interface DrawingPart {
  readonly drawingPath: string;
  readonly relationshipsPath: string;
  readonly drawingXml: string;
  readonly relationshipsXml: string;
  readonly media: readonly {
    readonly path: string;
    readonly bytes: Uint8Array;
    readonly mimeType: 'image/png' | 'image/jpeg';
  }[];
}

export interface DrawingMediaPool {
  readonly byHash: Map<
    string,
    {
      readonly path: string;
      readonly bytes: Uint8Array;
      readonly mimeType: 'image/png' | 'image/jpeg';
    }
  >;
}

export function createDrawingMediaPool(): DrawingMediaPool {
  return { byHash: new Map() };
}

function xml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function emu(value: number): number {
  return Math.round(value * EMU_PER_DIP);
}

function color(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const normalized = value.replace(/^#/, '');
  return /^[\dA-F]{6}$/i.test(normalized) ? normalized.toUpperCase() : undefined;
}

function marker(
  name: 'from' | 'to',
  value: {
    readonly row: number;
    readonly column: number;
    readonly offset: { readonly x: number; readonly y: number };
  },
): string {
  return `<xdr:${name}><xdr:col>${value.column}</xdr:col><xdr:colOff>${emu(value.offset.x)}</xdr:colOff><xdr:row>${value.row}</xdr:row><xdr:rowOff>${emu(value.offset.y)}</xdr:rowOff></xdr:${name}>`;
}

function anchor(object: SheetObject, content: string, clientData = '<xdr:clientData/>'): string {
  if (object.anchor.type === 'absolute') {
    const { rect } = object.anchor;
    return `<xdr:absoluteAnchor><xdr:pos x="${emu(rect.x)}" y="${emu(rect.y)}"/><xdr:ext cx="${emu(rect.width)}" cy="${emu(rect.height)}"/>${content}${clientData}</xdr:absoluteAnchor>`;
  }
  if (object.anchor.type === 'one-cell') {
    return `<xdr:oneCellAnchor>${marker('from', {
      ...object.anchor.cell,
      offset: object.anchor.offset,
    })}<xdr:ext cx="${emu(object.anchor.size.width)}" cy="${emu(object.anchor.size.height)}"/>${content}${clientData}</xdr:oneCellAnchor>`;
  }
  return `<xdr:twoCellAnchor editAs="twoCell">${marker('from', object.anchor.from)}${marker('to', object.anchor.to)}${content}${clientData}</xdr:twoCellAnchor>`;
}

function nonVisual(object: SheetObject, index: number, shape: boolean, textBox = false): string {
  const description =
    object.accessibility.description === undefined
      ? ''
      : ` descr="${xml(object.accessibility.description)}"`;
  const locks = object.locked
    ? `<a:${shape ? 'spLocks' : 'picLocks'} noMove="1" noResize="1"/>`
    : '';
  return shape
    ? `<xdr:nvSpPr><xdr:cNvPr id="${index + 1}" name="${xml(object.accessibility.name)}"${description}/><xdr:cNvSpPr${textBox ? ' txBox="1"' : ''}>${locks}</xdr:cNvSpPr></xdr:nvSpPr>`
    : `<xdr:nvPicPr><xdr:cNvPr id="${index + 1}" name="${xml(object.accessibility.name)}"${description}/><xdr:cNvPicPr>${locks}</xdr:cNvPicPr></xdr:nvPicPr>`;
}

function transform(object: SheetObject): string {
  return object.rotation === undefined || object.rotation === 0
    ? '<a:xfrm/>'
    : `<a:xfrm rot="${Math.round(object.rotation * 60_000)}"/>`;
}

function shapeProperties(object: Extract<SheetObject, { readonly kind: 'shape' }>): string {
  const fill = color(object.style.fill);
  const stroke = color(object.style.stroke);
  const line =
    stroke === undefined
      ? '<a:ln><a:noFill/></a:ln>'
      : `<a:ln${object.style.strokeWidth === undefined ? '' : ` w="${emu(object.style.strokeWidth)}"`}><a:solidFill><a:srgbClr val="${stroke}"/></a:solidFill></a:ln>`;
  return `<xdr:spPr>${transform(object)}<a:prstGeom prst="${object.shape === 'rectangle' ? 'rect' : object.shape}"><a:avLst/></a:prstGeom>${fill === undefined ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`}${line}</xdr:spPr>`;
}

function textBody(object: Extract<SheetObject, { readonly kind: 'text-box' }>): string {
  const textColor = color(object.style.color) ?? '000000';
  const alignment =
    object.style.horizontalAlign === undefined
      ? ''
      : ` algn="${object.style.horizontalAlign === 'center' ? 'ctr' : object.style.horizontalAlign === 'right' ? 'r' : 'l'}"`;
  return `<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr${alignment}/><a:r><a:rPr lang="en-US" sz="${Math.round(object.style.fontSize * 100)}"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:latin typeface="${xml(object.style.fontFamily)}"/></a:rPr><a:t>${xml(object.text)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></xdr:txBody>`;
}

function resourceFor(
  document: GeneratedDocument,
  object: Extract<SheetObject, { readonly kind: 'image' }>,
): ResolvedResource {
  const resource = document.resources.byReference[object.resourceId];
  if (
    resource === undefined ||
    (resource.mimeType !== 'image/png' && resource.mimeType !== 'image/jpeg')
  ) {
    throw outputError('XLSX_RESOURCE_UNSUPPORTED', `Object ${object.id} is not PNG or JPEG`, {
      details: { objectId: object.id, resourceId: object.resourceId },
    });
  }
  return resource;
}

function legacyObjects(sheet: Sheet, document: GeneratedDocument): readonly SheetObject[] {
  const known = new Set((sheet.objects ?? []).map(({ id }) => String(id)));
  return document.objects
    .filter(
      ({ generated, objectId, resourceId }) =>
        generated.sheetId === sheet.id && resourceId !== undefined && !known.has(String(objectId)),
    )
    .map(
      ({ objectId, resourceId, generated }): SheetObject =>
        ({
          id: objectId,
          kind: 'image',
          resourceId: resourceId!,
          anchor: {
            type: 'two-cell',
            from: {
              sheetId: sheet.id,
              row: generated.start.row,
              column: generated.start.column,
              offset: { x: 0, y: 0 },
            },
            to: {
              sheetId: sheet.id,
              row: generated.end.row + 1,
              column: generated.end.column + 1,
              offset: { x: 0, y: 0 },
            },
          },
          zIndex: 0,
          locked: false,
          templateRepeat: 'shared',
          accessibility: { name: String(objectId) },
        }) as SheetObject,
    );
}

export function createDrawingPart(
  sheet: Sheet,
  sheetIndex: number,
  document: GeneratedDocument,
  mediaPool: DrawingMediaPool = createDrawingMediaPool(),
): DrawingPart | undefined {
  const objects = [...(sheet.objects ?? []), ...legacyObjects(sheet, document)].sort(
    (left, right) =>
      left.zIndex - right.zIndex ||
      (String(left.id) < String(right.id) ? -1 : String(left.id) > String(right.id) ? 1 : 0),
  );
  if (objects.length === 0) return undefined;
  const createdMedia: DrawingPart['media'][number][] = [];
  const relationships: string[] = [];
  const anchors = objects.map((object, index) => {
    if (object.kind === 'image') {
      const resource = resourceFor(document, object);
      let media = mediaPool.byHash.get(resource.contentHash);
      if (media === undefined) {
        const mimeType: 'image/png' | 'image/jpeg' =
          resource.mimeType === 'image/png' ? 'image/png' : 'image/jpeg';
        const extension = mimeType === 'image/png' ? 'png' : 'jpeg';
        const created = {
          path: `xl/media/image${mediaPool.byHash.size + 1}.${extension}`,
          bytes: new Uint8Array(resource.bytes),
          mimeType,
        };
        mediaPool.byHash.set(resource.contentHash, created);
        createdMedia.push(created);
        media = created;
      }
      const relationshipId = `rId${relationships.length + 1}`;
      relationships.push(
        `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${media.path.split('/').at(-1)}"/>`,
      );
      const content = `<xdr:pic>${nonVisual(object, index, false)}<xdr:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${transform(object)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
      return anchor(object, content);
    }
    if (object.kind === 'shape') {
      return anchor(
        object,
        `<xdr:sp>${nonVisual(object, index, true)}${shapeProperties(object)}</xdr:sp>`,
      );
    }
    return anchor(
      object,
      `<xdr:sp>${nonVisual(object, index, true, true)}<xdr:spPr>${transform(object)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>${textBody(object)}</xdr:sp>`,
    );
  });
  return {
    drawingPath: `xl/drawings/drawing${sheetIndex + 1}.xml`,
    relationshipsPath: `xl/drawings/_rels/drawing${sheetIndex + 1}.xml.rels`,
    drawingXml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `${anchors.join('')}</xdr:wsDr>`,
    relationshipsXml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`,
    media: createdMedia,
  };
}
