import {
  createSpreadsheetDocument,
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  TegoSheetException,
  type CellInput,
  type DocumentParseResult,
  type LocaleDefinition,
  type Sheet,
  type SparseCell,
  type SpreadsheetDocument,
  type Workbook,
  type WorkbookSettings,
} from 'tego-sheet';
import { de } from 'tego-sheet/locales/de';
import { en } from 'tego-sheet/locales/en';
import { nl } from 'tego-sheet/locales/nl';
import { zhCN } from 'tego-sheet/locales/zh-cn';

const locales: readonly LocaleDefinition[] = [en, de, nl, zhCN];
const exception = new TegoSheetException({
  code: 'INVALID_COMMAND',
  message: 'Consumer type probe',
  recoverable: false,
});
const created: SpreadsheetDocument = createSpreadsheetDocument({
  id: 'consumer-document',
  sheetId: 'consumer-sheet',
});
const parsed: DocumentParseResult = parseSpreadsheetDocument(serializeSpreadsheetDocument(created));
const settings: WorkbookSettings = created.workbook.settings;
const workbook: Workbook = created.workbook;
const sheet: Sheet = workbook.sheets[0]!;
const sparseCell: SparseCell = {
  row: 0,
  column: 0,
  cell: { input: { type: 'blank' } },
};
const cellInput: CellInput = sparseCell.cell.input;

void locales;
void exception;
void parsed;
void settings;
void sheet;
void cellInput;
