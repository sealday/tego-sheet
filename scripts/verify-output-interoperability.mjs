import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { XlsxAdapter } from '../dist/output/xlsx.js';

const temporary = mkdtempSync(join(tmpdir(), 'tego-sheet-xlsx-interop-'));
const sourcePath = join(temporary, 'tego-sheet-source.xlsx');
const excelPath = join(temporary, 'tego-sheet-excel-roundtrip.xlsx');
const libreOfficeDirectory = join(temporary, 'libreoffice');
const libreOfficeExecutable = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
const excelExecutable = '/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel';

const document = {
  workbook: {
    sheets: [
      {
        id: 'invoice',
        name: 'Invoice',
        cells: [
          { row: 0, column: 0, cell: { input: { type: 'string', value: '=literal' } } },
          { row: 1, column: 0, cell: { input: { type: 'number', value: 2 } } },
          { row: 1, column: 1, cell: { input: { type: 'formula', source: '=A2+1' } } },
        ],
        merges: [],
        rows: [],
        columns: [],
      },
      {
        id: 'archive',
        name: 'Archive',
        cells: [],
        merges: [],
        rows: [],
        columns: [],
      },
    ],
    styles: [],
    validations: [],
    settings: { dateSystem: 'excel-1900' },
  },
  calculatedCells: [
    {
      address: { sheetId: 'invoice', row: 1, column: 1 },
      value: { type: 'number', value: 3 },
    },
  ],
  worksheets: [
    {
      sheetId: 'invoice',
      visibility: 'visible',
      conditionalFormatting: [
        {
          type: 'color-scale',
          range: {
            sheetId: 'invoice',
            start: { row: 1, column: 0 },
            end: { row: 1, column: 1 },
          },
          minimumColor: '#ff0000',
          maximumColor: '#00ff00',
        },
        {
          type: 'cell-is',
          range: {
            sheetId: 'invoice',
            start: { row: 1, column: 0 },
            end: { row: 1, column: 1 },
          },
          operator: 'greaterThan',
          formula: '1',
          style: { bold: true, backgroundColor: '#ffff00' },
        },
      ],
    },
    {
      sheetId: 'archive',
      visibility: 'very-hidden',
      conditionalFormatting: [],
    },
  ],
  print: {
    pages: [],
    displayList: { diagnostics: [], pages: [] },
    profile: {
      id: 'interop',
      name: 'Interop',
      targets: [
        {
          type: 'range',
          range: {
            sheetId: 'invoice',
            start: { row: 0, column: 0 },
            end: { row: 1, column: 1 },
          },
        },
      ],
      page: {
        paper: { type: 'custom', width: 384, height: 192 },
        orientation: 'landscape',
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        scale: { type: 'fixed', value: 1 },
      },
      manualBreaks: [],
      showGridlines: true,
      showHeadings: true,
    },
  },
  resources: { byHash: {}, byReference: {}, totalBytes: 0 },
  objects: [],
  diagnostics: [],
  metadata: {
    templateId: 'interop',
    profileId: 'interop',
    sourceDocumentHash: 'sha256:interop',
    locale: 'en-US',
    timeZone: 'UTC',
    generatedAt: '2026-07-23T00:00:00.000Z',
  },
};

const pythonProbe = String.raw`
import json, re, sys
from openpyxl import load_workbook
def inches(value):
    match = re.fullmatch(r"([0-9.]+)(in|mm)", str(value))
    assert match, value
    number, unit = float(match.group(1)), match.group(2)
    return number if unit == "in" else number / 25.4
book = load_workbook(sys.argv[1], data_only=False)
sheet = book["Invoice"]
expected_archive_state = sys.argv[2]
assert book["Archive"].sheet_state == expected_archive_state, book["Archive"].sheet_state
assert sheet["A1"].value == "=literal" and sheet["A1"].data_type == "s", (sheet["A1"].value, sheet["A1"].data_type)
assert sheet["B2"].value == "=A2+1" and sheet["B2"].data_type == "f", (sheet["B2"].value, sheet["B2"].data_type)
assert str(sheet.print_area) == "'Invoice'!$A$1:$B$2", str(sheet.print_area)
assert abs(inches(sheet.page_setup.paperWidth) - 4) < 0.02, str(sheet.page_setup.paperWidth)
assert abs(inches(sheet.page_setup.paperHeight) - 2) < 0.02, str(sheet.page_setup.paperHeight)
conditional_rules = [rule for rules in sheet.conditional_formatting._cf_rules.values() for rule in rules]
assert sorted(rule.type for rule in conditional_rules) == ["cellIs", "colorScale"], [rule.type for rule in conditional_rules]
print(json.dumps({"openpyxl": "pass", "sheets": book.sheetnames}))
`;

const excelProbe = String.raw`
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  tell application "Microsoft Excel"
    set wb to open workbook workbook file name inputPath read only true
    set sheetCount to count of worksheets of wb
    set formulaValue to formula of range "B2" of worksheet 1 of wb
    save workbook as wb filename outputPath file format Excel XML file format
    close wb saving no
    return (sheetCount as text) & "|" & formulaValue
  end tell
end run
`;

function processIds(name) {
  try {
    return new Set(
      execFileSync('pgrep', ['-x', name], { encoding: 'utf8' })
        .trim()
        .split('\n')
        .filter(Boolean)
        .map(Number),
    );
  } catch {
    return new Set();
  }
}

async function closeNewProcesses(name, previous) {
  const current = processIds(name);
  const created = [...current].filter((pid) => !previous.has(pid));
  for (const pid of created) process.kill(pid, 'SIGTERM');
  if (created.length === 0) return;
  await new Promise((resolve) => setTimeout(resolve, 5_000));
  for (const pid of created) {
    try {
      process.kill(pid, 0);
      process.kill(pid, 'SIGKILL');
    } catch {
      // The process exited after SIGTERM.
    }
  }
}

try {
  const blob = await new XlsxAdapter().render(document, {
    formulaMode: 'formula-and-cached-value',
    compatibility: 'excel',
  });
  writeFileSync(sourcePath, new Uint8Array(await blob.arrayBuffer()));

  const openpyxl = execFileSync('python3', ['-c', pythonProbe, sourcePath, 'veryHidden'], {
    encoding: 'utf8',
  }).trim();

  let libreOffice = {
    status: 'not-executed',
    reason: 'LibreOffice is not installed at the expected macOS application path.',
  };
  if (existsSync(libreOfficeExecutable)) {
    mkdirSync(libreOfficeDirectory);
    const output = execFileSync(
      libreOfficeExecutable,
      ['--headless', '--convert-to', 'xlsx', '--outdir', libreOfficeDirectory, sourcePath],
      { encoding: 'utf8' },
    ).trim();
    const libreOfficePath = join(libreOfficeDirectory, 'tego-sheet-source.xlsx');
    readFileSync(libreOfficePath);
    execFileSync('python3', ['-c', pythonProbe, libreOfficePath, 'hidden'], {
      encoding: 'utf8',
    });
    libreOffice = {
      status: 'pass',
      output,
      normalization: [
        'veryHidden worksheet state is downgraded to hidden',
        'inch paper dimensions are converted to equivalent millimetres',
      ],
    };
  }

  let excelDesktop = {
    status: 'not-executed',
    reason:
      process.env.TEGO_SHEET_EXCEL_INTEROP === '1'
        ? 'Microsoft Excel is not installed at the expected macOS application path.'
        : 'Set TEGO_SHEET_EXCEL_INTEROP=1 to run the opt-in desktop automation probe.',
  };
  if (process.env.TEGO_SHEET_EXCEL_INTEROP === '1' && existsSync(excelExecutable)) {
    const previousExcelProcesses = processIds('Microsoft Excel');
    try {
      const result = execFileSync('osascript', ['-', sourcePath, excelPath], {
        encoding: 'utf8',
        input: excelProbe,
        timeout: 60_000,
      }).trim();
      readFileSync(excelPath);
      execFileSync('python3', ['-c', pythonProbe, excelPath, 'veryHidden'], {
        encoding: 'utf8',
      });
      excelDesktop = { status: 'pass', result };
    } catch (error) {
      excelDesktop = {
        status: 'blocked',
        reason:
          error instanceof Error ? `${error.name}: ${error.message.split('\n')[0]}` : String(error),
      };
      await closeNewProcesses('Microsoft Excel', previousExcelProcesses);
    }
  }

  console.log(
    JSON.stringify(
      {
        implementationGate: 'pass',
        externalReleaseGate: 'pending',
        openpyxl: JSON.parse(openpyxl),
        libreOffice,
        excelDesktop,
        excelWeb: {
          status: 'not-executed',
          reason:
            'No authorized Excel Web upload/account channel is available in this environment.',
        },
      },
      null,
      2,
    ),
  );
  process.exitCode = 2;
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
