import type { WorkbookData } from 'tego-sheet';
import type { PlaygroundMode } from './playground-model';

export interface PlaygroundPreset {
  readonly mode: PlaygroundMode;
  readonly label: string;
  readonly description: string;
  readonly docsLink: string;
  readonly publicApis: readonly string[];
  readonly createFixture: () => WorkbookData;
}

export const PLAYGROUND_LOCALES = Object.freeze([
  Object.freeze({ label: 'English', subpath: 'tego-sheet/locales/en' }),
  Object.freeze({ label: '简体中文', subpath: 'tego-sheet/locales/zh-cn' }),
  Object.freeze({ label: 'Deutsch', subpath: 'tego-sheet/locales/de' }),
  Object.freeze({ label: 'Nederlands', subpath: 'tego-sheet/locales/nl' }),
] as const);

export function createUncontrolledFixture(): WorkbookData {
  return [
    {
      name: 'Budget',
      freeze: 'B2',
      rows: {
        len: 20,
        0: { cells: { 0: { text: 'Item' }, 1: { text: 'Amount' } } },
        1: { cells: { 0: { text: 'Hosting' }, 1: { text: '29' } } },
        2: { cells: { 0: { text: 'Support' }, 1: { text: '75' } } },
        3: { cells: { 0: { text: 'Total' }, 1: { text: '=SUM(B2:B3)' } } },
      },
      cols: { len: 6 },
    },
  ];
}

export function createControlledFixture(): WorkbookData {
  return [
    {
      name: 'Inventory',
      rows: {
        len: 20,
        0: { cells: { 0: { text: 'Product' }, 1: { text: 'Quantity' } } },
        1: { cells: { 0: { text: 'Keyboard' }, 1: { text: '12' } } },
        2: { cells: { 0: { text: 'Mouse' }, 1: { text: '24' } } },
      },
      cols: { len: 6 },
    },
  ];
}

export function createCustomChromeFixture(): WorkbookData {
  return [
    {
      name: 'Roadmap',
      rows: {
        len: 20,
        0: { cells: { 0: { text: 'Milestone' }, 1: { text: 'Status' } } },
        1: { cells: { 0: { text: 'Prototype' }, 1: { text: 'Complete' } } },
        2: { cells: { 0: { text: 'Documentation' }, 1: { text: 'In progress' } } },
      },
      cols: { len: 6 },
    },
  ];
}

export function createLocalesFixture(): WorkbookData {
  return [
    {
      name: 'Locale demo',
      rows: {
        len: 20,
        0: { cells: { 0: { text: 'Language' }, 1: { text: 'Greeting' } } },
        1: { cells: { 0: { text: 'English' }, 1: { text: 'Hello' } } },
        2: { cells: { 0: { text: '简体中文' }, 1: { text: '你好' } } },
        3: { cells: { 0: { text: 'Deutsch' }, 1: { text: 'Hallo' } } },
        4: { cells: { 0: { text: 'Nederlands' }, 1: { text: 'Hallo' } } },
      },
      cols: { len: 6 },
    },
  ];
}

export function createLegacyJsonFixture(): WorkbookData {
  return [
    {
      name: '',
      freeze: 'A1',
      styles: [{ strike: false, textwrap: false, underline: false }],
      rows: {
        len: 80,
        0: {
          height: 0,
          hide: false,
          style: 0,
          cells: {
            0: {
              text: '',
              style: 0,
              editable: false,
              printable: false,
              value: 0,
            },
            12: { text: 'edge' },
          },
        },
        47: { hide: true, cells: { 31: { text: 'sparse' } } },
      },
      cols: {
        len: 50,
        0: { width: 0, hide: false, style: 0 },
        23: { width: 64, hide: true },
      },
    },
  ];
}

export const PLAYGROUND_PRESETS = {
  uncontrolled: {
    mode: 'uncontrolled',
    label: 'Uncontrolled',
    description: 'Let TegoSheet own edits after reading defaultValue once at mount.',
    docsLink: '/docs/concepts/controlled-and-uncontrolled',
    publicApis: ['TegoSheet', 'defaultValue', 'onChange'],
    createFixture: createUncontrolledFixture,
  },
  controlled: {
    mode: 'controlled',
    label: 'Controlled',
    description: 'Accept each onChange snapshot into a parent-owned value.',
    docsLink: '/docs/concepts/controlled-and-uncontrolled',
    publicApis: ['TegoSheet', 'value', 'onChange'],
    createFixture: createControlledFixture,
  },
  'custom-chrome': {
    mode: 'custom-chrome',
    label: 'Custom Chrome',
    description: 'Replace built-in chrome with typed toolbar and sheet-tab renderers.',
    docsLink: '/docs/guides/custom-chrome',
    publicApis: ['TegoSheet', 'toolbar', 'sheetTabs'],
    createFixture: createCustomChromeFixture,
  },
  locales: {
    mode: 'locales',
    label: 'Locales',
    description: 'Switch one spreadsheet among the four published locale dictionaries.',
    docsLink: '/docs/guides/locales',
    publicApis: ['TegoSheet', 'locale'],
    createFixture: createLocalesFixture,
  },
  'legacy-json': {
    mode: 'legacy-json',
    label: 'Legacy JSON',
    description: 'Load compatible sparse workbook JSON and inspect its public snapshot.',
    docsLink: '/docs/migration/from-x-data-spreadsheet',
    publicApis: ['TegoSheet', 'WorkbookInput', 'TegoSheetHandle.getValue'],
    createFixture: createLegacyJsonFixture,
  },
} as const satisfies { readonly [Mode in PlaygroundMode]: PlaygroundPreset };

export function createFixture(mode: PlaygroundMode): WorkbookData {
  return PLAYGROUND_PRESETS[mode].createFixture();
}
