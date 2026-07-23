import {
  migrateLegacyWorkbook,
  parseSpreadsheetDocument,
  type SpreadsheetDocument,
} from 'tego-sheet';

export const PREVIEW_EVENT_LIMIT = 12;

export interface PreviewEventInput {
  readonly timestamp: string;
  readonly label: string;
  readonly details?: string;
}

export interface PreviewEvent extends PreviewEventInput {
  readonly id: string;
}

const EXAMPLE_WORKBOOK = [
  {
    name: 'Budget',
    freeze: 'B2',
    rows: {
      len: 5,
      0: { cells: { 0: { text: 'Item' }, 1: { text: 'Amount' } } },
      1: { cells: { 0: { text: 'Hosting' }, 1: { text: '29' } } },
      2: { cells: { 0: { text: 'Support' }, 1: { text: '75' } } },
      3: { cells: { 0: { text: 'Total' }, 1: { text: '=SUM(B2:B3)' } } },
    },
    cols: { len: 4 },
  },
] as const;

function isSheetData(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function invalidKnownField(index: number, field: string, expectation: string): never {
  throw new TypeError(`Workbook data is invalid: workbook[${index}].${field} ${expectation}.`);
}

function validateKnownSheetFields(sheet: Record<string, unknown>, index: number): void {
  if (sheet.name !== undefined && typeof sheet.name !== 'string') {
    invalidKnownField(index, 'name', 'must be a string');
  }
  if (sheet.freeze !== undefined && typeof sheet.freeze !== 'string') {
    invalidKnownField(index, 'freeze', 'must be a string');
  }
  for (const field of ['rows', 'cols'] as const) {
    const value = sheet[field];
    if (value !== undefined && !isSheetData(value)) {
      invalidKnownField(index, field, 'must be a JSON object');
    }
  }
}

function createEventId(logs: readonly PreviewEvent[], timestamp: string): string {
  const baseId = `preview-event-${timestamp}`;
  let id = baseId;
  let suffix = 1;

  while (logs.some((entry) => entry.id === id)) {
    id = `${baseId}-${suffix}`;
    suffix += 1;
  }

  return id;
}

function migrate(input: unknown): SpreadsheetDocument {
  const result = migrateLegacyWorkbook(input);
  if (!result.ok) throw new TypeError('Workbook JSON could not be migrated to schema 2.');
  return result.document;
}

export function cloneExampleWorkbook(): SpreadsheetDocument {
  return migrate(EXAMPLE_WORKBOOK);
}

export function parseWorkbookJson(source: string): SpreadsheetDocument {
  const parsed: unknown = JSON.parse(source);
  const schema = parseSpreadsheetDocument(parsed);
  if (schema.ok) return schema.document;

  if (Array.isArray(parsed)) {
    if (!parsed.every(isSheetData)) {
      throw new TypeError('Workbook JSON must be a sheet object or an array of sheet objects.');
    }

    parsed.forEach(validateKnownSheetFields);
    return migrate(parsed);
  }

  if (!isSheetData(parsed)) {
    throw new TypeError('Workbook JSON must be a sheet object or an array of sheet objects.');
  }

  validateKnownSheetFields(parsed, 0);
  return migrate(parsed);
}

export function formatWorkbookJson(document: SpreadsheetDocument): string {
  return JSON.stringify(document, null, 2);
}

export function appendPreviewEvent(
  logs: readonly PreviewEvent[],
  input: Readonly<PreviewEventInput>,
): PreviewEvent[] {
  const entry: PreviewEvent = {
    ...input,
    id: createEventId(logs, input.timestamp),
  };

  return [entry, ...logs].slice(0, PREVIEW_EVENT_LIMIT);
}
