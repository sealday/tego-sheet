import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, expectTypeOf, it } from 'vitest';
import type { SpreadsheetDocument } from 'tego-sheet';
import type { WorkbookData } from '../../../src/core';
import {
  PLAYGROUND_MODES,
  appendPlaygroundEvent,
  parsePlaygroundMode,
  type PlaygroundEvent,
} from '../../../website/src/components/playground/playground-model';
import {
  PLAYGROUND_LOCALES,
  PLAYGROUND_PRESETS,
  createControlledFixture,
  createCustomChromeFixture,
  createFixture,
  createLegacyJsonFixture,
  createLocalesFixture,
  createUncontrolledFixture,
} from '../../../website/src/components/playground/playground-fixtures';

describe('playground modes', () => {
  it('keeps the five URL modes in the approved order', () => {
    expect(PLAYGROUND_MODES).toEqual([
      'uncontrolled',
      'controlled',
      'custom-chrome',
      'locales',
      'legacy-json',
    ]);
  });

  it.each([
    [null, 'uncontrolled'],
    ['controlled', 'controlled'],
    ['private-engine', 'uncontrolled'],
  ] as const)('parses %s as %s', (input, expected) => {
    expect(parsePlaygroundMode(input)).toBe(expected);
  });
});

describe('playground event history', () => {
  it('retains only the newest 50 records without mutating its input', () => {
    const existing = Array.from(
      { length: 60 },
      (_, sequence): PlaygroundEvent => ({
        sequence,
        callback: 'onSelectionChange',
        payload: { sequence },
      }),
    );
    const before = structuredClone(existing);

    const result = appendPlaygroundEvent(existing, {
      sequence: 60,
      callback: 'onDocumentChange',
      payload: { sequence: 60 },
    });

    expect(result).toHaveLength(50);
    expect(result.map(({ sequence }) => sequence)).toEqual(
      Array.from({ length: 50 }, (_, index) => index + 11),
    );
    expect(existing).toEqual(before);
    expect(existing).toHaveLength(60);
  });

  it('returns deeply isolated immutable records with JSON-serializable payloads', () => {
    const payload = { nested: { value: 'original' }, entries: [1, false, null] };
    const input: PlaygroundEvent[] = [
      { sequence: 1, callback: 'onCellEdit', payload: { retained: { value: 'input' } } },
    ];

    const result = appendPlaygroundEvent(input, {
      sequence: 2,
      callback: 'onError',
      payload,
    });
    payload.nested.value = 'mutated later';
    const appended = result[1];
    if (!appended) throw new Error('the appended event must be retained');

    expect(result).not.toBe(input);
    expect(result[0]).not.toBe(input[0]);
    expect(appended.payload).toEqual({
      nested: { value: 'original' },
      entries: [1, false, null],
    });
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result[0])).toBe(true);
    expect(Object.isFrozen(result[0]?.payload)).toBe(true);
    expect(Object.isFrozen(appended.payload)).toBe(true);
    expect(Object.isFrozen((appended.payload as { readonly nested: object }).nested)).toBe(true);
  });

  it('preserves an own __proto__ JSON key without changing object prototypes', () => {
    const payload = JSON.parse('{"__proto__":{"polluted":true}}') as {
      readonly __proto__: { readonly polluted: boolean };
    };
    const [event] = appendPlaygroundEvent([], {
      sequence: 1,
      callback: 'onDocumentChange',
      payload,
    });
    if (!event || typeof event.payload !== 'object' || event.payload === null) {
      throw new Error('the appended object payload must be retained');
    }

    expect(Object.hasOwn(event.payload, '__proto__')).toBe(true);
    expect(Object.getPrototypeOf(event.payload)).toBe(Object.prototype);
    expect(JSON.stringify(event.payload)).toBe('{"__proto__":{"polluted":true}}');
    expect(({} as { polluted?: boolean }).polluted).toBeUndefined();
    expect(Object.isFrozen(event.payload)).toBe(true);
  });
});

describe('playground preset registry', () => {
  it('publishes exact labels, guide links, descriptions, and public API names', () => {
    expect(
      PLAYGROUND_MODES.map((mode) => {
        const preset = PLAYGROUND_PRESETS[mode];
        return {
          mode: preset.mode,
          label: preset.label,
          description: preset.description,
          docsLink: preset.docsLink,
          publicApis: preset.publicApis,
        };
      }),
    ).toEqual([
      {
        mode: 'uncontrolled',
        label: 'Uncontrolled',
        description: 'Let TegoSheet own edits after reading defaultDocument once at mount.',
        docsLink: '/docs/concepts/controlled-and-uncontrolled',
        publicApis: ['TegoSheet', 'defaultDocument', 'onDocumentChange'],
      },
      {
        mode: 'controlled',
        label: 'Controlled',
        description: 'Accept each document snapshot into parent-owned state.',
        docsLink: '/docs/concepts/controlled-and-uncontrolled',
        publicApis: ['TegoSheet', 'document', 'onDocumentChange'],
      },
      {
        mode: 'custom-chrome',
        label: 'Custom Chrome',
        description: 'Replace built-in chrome with typed toolbar and sheet-tab renderers.',
        docsLink: '/docs/guides/custom-chrome',
        publicApis: ['TegoSheet', 'toolbar', 'sheetTabs'],
      },
      {
        mode: 'locales',
        label: 'Locales',
        description: 'Switch one spreadsheet among the four published locale dictionaries.',
        docsLink: '/docs/guides/locales',
        publicApis: ['TegoSheet', 'locale'],
      },
      {
        mode: 'legacy-json',
        label: 'Legacy JSON',
        description: 'Load compatible sparse workbook JSON and inspect its public snapshot.',
        docsLink: '/docs/migration/from-x-data-spreadsheet',
        publicApis: ['migrateLegacyWorkbook', 'TegoSheet', 'TegoSheetHandle.getDocument'],
      },
    ]);

    expect(Object.isFrozen(PLAYGROUND_PRESETS)).toBe(true);
    for (const preset of Object.values(PLAYGROUND_PRESETS)) {
      expect(Object.isFrozen(preset)).toBe(true);
      expect(Object.isFrozen(preset.publicApis)).toBe(true);
    }
  });

  it('maps locales only to the four published package subpaths', () => {
    expect(PLAYGROUND_LOCALES).toEqual([
      { label: 'English', subpath: 'tego-sheet/locales/en' },
      { label: '简体中文', subpath: 'tego-sheet/locales/zh-cn' },
      { label: 'Deutsch', subpath: 'tego-sheet/locales/de' },
      { label: 'Nederlands', subpath: 'tego-sheet/locales/nl' },
    ]);
  });

  it('exposes five independent fixture factories', () => {
    const factories = [
      createUncontrolledFixture,
      createControlledFixture,
      createCustomChromeFixture,
      createLocalesFixture,
      createLegacyJsonFixture,
    ];

    for (const [index, mode] of PLAYGROUND_MODES.entries()) {
      const first = factories[index]!();
      const second = createFixture(mode);

      expect(first).toBeDefined();
      expect(second.schemaVersion).toBe(2);
      expect(Object.isFrozen(second)).toBe(true);
      expectTypeOf(second).toMatchTypeOf<SpreadsheetDocument>();
    }
  });

  it('does not leak nested fixture mutations across calls', () => {
    const first = createFixture('legacy-json');
    const second = createFixture('legacy-json');
    expect(first).not.toBe(second);
    expect(first.workbook.sheets[0]?.cells).not.toBe(second.workbook.sheets[0]?.cells);
    expect(Object.isFrozen(first.workbook.sheets[0]?.cells)).toBe(true);
  });

  it('uses the established compatible sparse legacy JSON shape', () => {
    const fixturePath = fileURLToPath(
      new URL('../../parity/fixtures/workbooks/sparse-falsy.json', import.meta.url),
    );
    const canonical = JSON.parse(readFileSync(fixturePath, 'utf8')) as {
      readonly input: WorkbookData[number];
    };

    expect(createLegacyJsonFixture()).toEqual([canonical.input]);
  });
});
