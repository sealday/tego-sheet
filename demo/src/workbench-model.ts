import type { WorkbookData, WorkbookInput } from 'tego-sheet';

export const PREVIEW_EVENT_LIMIT = 12;

export interface PreviewEventInput {
  readonly timestamp: string;
  readonly label: string;
  readonly details?: string;
}

export interface PreviewEvent extends PreviewEventInput {
  readonly id: string;
}

const EXAMPLE_WORKBOOK: WorkbookData = [{
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
}];

function isSheetData(value: unknown): value is WorkbookData[number] {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function invalidKnownField(index: number, field: string, expectation: string): never {
  throw new TypeError(`Workbook data is invalid: workbook[${index}].${field} ${expectation}.`);
}

function validateKnownSheetFields(sheet: WorkbookData[number], index: number): void {
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

  while (logs.some(entry => entry.id === id)) {
    id = `${baseId}-${suffix}`;
    suffix += 1;
  }

  return id;
}

export function cloneExampleWorkbook(): WorkbookData {
  return JSON.parse(JSON.stringify(EXAMPLE_WORKBOOK)) as WorkbookData;
}

export function parseWorkbookJson(source: string): WorkbookInput {
  const parsed: unknown = JSON.parse(source);

  if (Array.isArray(parsed)) {
    if (!parsed.every(isSheetData)) {
      throw new TypeError('Workbook JSON must be a sheet object or an array of sheet objects.');
    }

    parsed.forEach(validateKnownSheetFields);
    return parsed;
  }

  if (!isSheetData(parsed)) {
    throw new TypeError('Workbook JSON must be a sheet object or an array of sheet objects.');
  }

  validateKnownSheetFields(parsed, 0);
  return parsed;
}

export function formatWorkbookJson(workbook: WorkbookInput): string {
  return JSON.stringify(workbook, null, 2);
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
